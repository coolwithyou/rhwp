import test from 'node:test';
import assert from 'node:assert/strict';

import { parseApplyFieldCommand } from '../src/document-agent/contract.ts';

// 5차: restore-exact 영역 문단의 선택 필드 endCharShapeId(field-end-char-shape-v1) 파싱.

function restoreCommand(paragraphExtra: Record<string, unknown> = {}) {
  return {
    schemaVersion: 1,
    commandId: 'region-restore-1',
    expectedDocumentEpoch: 1,
    expectedChangeSeq: 0,
    expectedDocumentSha256: 'a'.repeat(64),
    target: { kind: 'table_cell_region', section: 0, parentPara: 13, controlIndex: 0, cellIndex: 4 },
    expectedBeforeSha256: 'b'.repeat(64),
    expectedFormatSha256: 'c'.repeat(64),
    expectedAdjacentContextSha256: 'd'.repeat(64),
    replacement: '첫문단\n둘',
    replacementStyle: 'restore-exact',
    replacementFormat: {
      kind: 'table_cell_region',
      paragraphs: [
        { length: 3, charShapeIds: [70, 70, 70], paraShapeId: 28, ...paragraphExtra },
        { length: 1, charShapeIds: [70], paraShapeId: 28 },
      ],
    },
    expectedReplacementFormatSha256: 'e'.repeat(64),
  };
}

test('endCharShapeId가 없으면 이전 계약 그대로 파싱한다', () => {
  const parsed = parseApplyFieldCommand(restoreCommand());
  assert.equal(parsed.replacementFormat?.kind, 'table_cell_region');
  assert.ok(parsed.replacementFormat?.kind === 'table_cell_region'
    && parsed.replacementFormat.paragraphs.every(paragraph => !('endCharShapeId' in paragraph)));
});

test('endCharShapeId는 문단별 선택 필드로 보존한다', () => {
  const parsed = parseApplyFieldCommand(restoreCommand({ endCharShapeId: 21 }));
  assert.ok(parsed.replacementFormat?.kind === 'table_cell_region');
  assert.equal(parsed.replacementFormat.paragraphs[0]!.endCharShapeId, 21);
  assert.equal(parsed.replacementFormat.paragraphs[1]!.endCharShapeId, undefined);
});

test('잘못된 endCharShapeId와 다른 extra key는 거부한다', () => {
  for (const extra of [{ endCharShapeId: -1 }, { endCharShapeId: 2.5 }, { endCharShapeId: '21' }, { trailingRun: 21 }]) {
    assert.throws(() => parseApplyFieldCommand(restoreCommand(extra)), (error: unknown) =>
      (error as { code?: string }).code === 'INVALID_COMMAND');
  }
});
