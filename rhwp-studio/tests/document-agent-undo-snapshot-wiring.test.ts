import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

// 5차: 같은 세션 field revert가 apply SnapshotCommand의 before 스냅숏을 쓰는 배선 가드.
// engine 클래스는 strip-only TS로 실행할 수 없어(parameter property) 이 저장소 관례대로
// 소스 배선을 검증하고, 행위는 controller fake 테스트와 실제 브라우저 검증으로 확인한다.

const rootDir = dirname(dirname(fileURLToPath(import.meta.url)));
const source = (rel: string): string => readFileSync(join(rootDir, rel), 'utf8');

function block(src: string, signature: string): string {
  const start = src.indexOf(signature);
  assert.notEqual(start, -1, `${signature} not found`);
  const next = src.slice(start + signature.length)
    .search(/\n {2}(?:public |private |protected |static |async |get |set )*[a-zA-Z][\w]*\s*\(/);
  return next === -1 ? src.slice(start) : src.slice(start, start + signature.length + next);
}

const commandFull = source('src/engine/command.ts');
const snapStart = commandFull.indexOf('export class SnapshotCommand');
const snapEnd = commandFull.indexOf('\nexport class ', snapStart + 1);
const snapshotCommand = commandFull.slice(snapStart, snapEnd === -1 ? undefined : snapEnd);
const inputHandler = source('src/engine/input-handler.ts');

test('SnapshotCommand는 before 스냅숏 id를 읽기 전용으로 노출하고 discard 뒤에는 null이다', () => {
  const getter = block(snapshotCommand, 'beforeSnapshotId(): number | null {');
  assert.match(getter, /return this\.beforeId;/);
  const discard = block(snapshotCommand, 'discard(wasm: WasmBridge): void {');
  assert.match(discard, /this\.beforeId = null;/, 'discard는 before id를 비워 hook이 null을 주게 한다');
  assert.match(block(snapshotCommand, 'cursorBeforePosition(): DocumentPosition {'), /\{ \.\.\.this\.cursorBefore \}/);
});

test('InputHandler hook은 undo top의 같은 종류 document-agent snapshot만 인정한다', () => {
  const hook = block(inputHandler, 'documentAgentUndoSnapshot(');
  assert.match(hook, /this\.history\.peekUndoTop\(\)/, 'undo top만 본다(스택 아래 명령 사용 금지)');
  assert.match(hook, /instanceof SnapshotCommand/);
  assert.match(hook, /top\.type !== `snapshot:\$\{operationType\}`/);
  assert.match(hook, /if \(snapshotId === null\) return null;/, '축출·discard된 명령은 null');
  assert.doesNotMatch(hook, /restoreSnapshot|saveSnapshot|discardSnapshot|history\.(undo|redo|execute)/,
    'hook은 읽기 전용이다');
});

test('controller는 같은 종류 hook 결과와 apply 때 기록한 id가 같을 때만 스냅숏 복원을 시도한다', () => {
  const controller = source('src/document-agent/controller.ts');
  const revert = block(controller, 'async revertFieldCommand(');
  assert.match(revert, /documentAgentUndoSnapshot\?\.\('field-agent:apply'\)/);
  assert.match(revert, /undoSnapshot\.snapshotId === entry\.applySnapshotId/);
  assert.match(revert, /exportDocumentSha256\(wasm, state\.format\) !== entry\.applyReceipt\.beforeDocumentSha256/,
    '복원 bytes가 apply 전 문서와 같아야 한다');
  // v1 보존 가드는 두 경로 공통으로 유지한다.
  for (const guard of [
    'COMMAND_NOT_LATEST',
    'DOCUMENT_EPOCH_MISMATCH',
    'CHANGE_SEQ_MISMATCH',
    'DOCUMENT_SHA_MISMATCH',
    "revertedEvidence.formatSha256 !== entry.beforeEvidence.formatSha256",
    'PAGE_COUNT_CHANGED',
  ]) assert.ok(revert.includes(guard), `${guard} 유지`);
  const apply = block(controller, 'async applyFieldCommand(');
  assert.match(apply, /applySnapshotId: this\.deps\.input\.documentAgentUndoSnapshot\?\.\('field-agent:apply'\)\?\.snapshotId \?\? null/);
});
