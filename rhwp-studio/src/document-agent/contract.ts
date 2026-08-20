import {
  DocumentAgentError,
  type RhwpApplyFieldCommandV1,
  type RhwpApplyTextCommandV1,
  type RhwpBodyParagraphTargetV1,
  type RhwpFieldTargetV1,
  type RhwpFormTextTargetV1,
  type RhwpRevertFieldCommandV1,
  type RhwpRevertTextCommandV1,
  type RhwpTableCellRegionTargetV1,
  type RhwpTableCellTextTargetV1,
} from './types.ts';

const SHA256 = /^[0-9a-f]{64}$/;

function record(value: unknown, label: string): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new DocumentAgentError('INVALID_COMMAND', `${label}은 객체여야 합니다.`);
  }
  return value as Record<string, unknown>;
}

function exactKeys(value: Record<string, unknown>, allowed: readonly string[], label: string): void {
  const actual = Object.keys(value).sort();
  const expected = [...allowed].sort();
  if (actual.length !== expected.length
      || actual.some((key, index) => key !== expected[index])) {
    throw new DocumentAgentError('INVALID_COMMAND', `${label} 필드가 strict schema와 다릅니다.`);
  }
}

function safeInteger(value: unknown, minimum: number, label: string): asserts value is number {
  if (!Number.isSafeInteger(value) || (value as number) < minimum) {
    throw new DocumentAgentError('INVALID_COMMAND', `${label}은 ${minimum} 이상의 safe integer여야 합니다.`);
  }
}

function digest(value: unknown, label: string): asserts value is string {
  if (typeof value !== 'string' || !SHA256.test(value)) {
    throw new DocumentAgentError('INVALID_COMMAND', `${label}은 lowercase SHA-256 hex여야 합니다.`);
  }
}

function commandId(value: unknown): asserts value is string {
  if (typeof value !== 'string' || value.length < 1 || value.length > 128) {
    throw new DocumentAgentError('INVALID_COMMAND', 'commandId 길이는 1..=128이어야 합니다.');
  }
}

export function parseBodyParagraphTarget(value: unknown): RhwpBodyParagraphTargetV1 {
  const target = record(value, 'target');
  exactKeys(target, ['kind', 'section', 'paragraph', 'charOffset', 'length'], 'target');
  if (target.kind !== 'body_paragraph') {
    throw new DocumentAgentError('INVALID_COMMAND', 'target.kind는 body_paragraph여야 합니다.');
  }
  safeInteger(target.section, 0, 'target.section');
  safeInteger(target.paragraph, 0, 'target.paragraph');
  if (target.charOffset !== 0) {
    throw new DocumentAgentError('INVALID_COMMAND', 'target.charOffset은 0이어야 합니다.');
  }
  safeInteger(target.length, 0, 'target.length');
  if (target.length > 4000) {
    throw new DocumentAgentError('INVALID_COMMAND', 'target.length는 4000 이하여야 합니다.');
  }
  return target as unknown as RhwpBodyParagraphTargetV1;
}

export function parseTableCellTextTarget(value: unknown): RhwpTableCellTextTargetV1 {
  const target = record(value, 'target');
  exactKeys(
    target,
    ['kind', 'section', 'parentPara', 'controlIndex', 'cellIndex', 'cellParagraph'],
    'target',
  );
  if (target.kind !== 'table_cell_text') {
    throw new DocumentAgentError('INVALID_COMMAND', 'target.kind는 table_cell_text여야 합니다.');
  }
  safeInteger(target.section, 0, 'target.section');
  safeInteger(target.parentPara, 0, 'target.parentPara');
  safeInteger(target.controlIndex, 0, 'target.controlIndex');
  safeInteger(target.cellIndex, 0, 'target.cellIndex');
  safeInteger(target.cellParagraph, 0, 'target.cellParagraph');
  return target as unknown as RhwpTableCellTextTargetV1;
}

export function parseTableCellRegionTarget(value: unknown): RhwpTableCellRegionTargetV1 {
  const target = record(value, 'target');
  exactKeys(target, ['kind', 'section', 'parentPara', 'controlIndex', 'cellIndex'], 'target');
  if (target.kind !== 'table_cell_region') {
    throw new DocumentAgentError('INVALID_COMMAND', 'target.kind는 table_cell_region이어야 합니다.');
  }
  safeInteger(target.section, 0, 'target.section');
  safeInteger(target.parentPara, 0, 'target.parentPara');
  safeInteger(target.controlIndex, 0, 'target.controlIndex');
  safeInteger(target.cellIndex, 0, 'target.cellIndex');
  return target as unknown as RhwpTableCellRegionTargetV1;
}

export function parseFormTextTarget(value: unknown): RhwpFormTextTargetV1 {
  const target = record(value, 'target');
  exactKeys(target, ['kind', 'section', 'paragraph', 'fieldId'], 'target');
  if (target.kind !== 'form_text') {
    throw new DocumentAgentError('INVALID_COMMAND', 'target.kind는 form_text여야 합니다.');
  }
  safeInteger(target.section, 0, 'target.section');
  safeInteger(target.paragraph, 0, 'target.paragraph');
  safeInteger(target.fieldId, 0, 'target.fieldId');
  return target as unknown as RhwpFormTextTargetV1;
}

export function parseFieldTarget(value: unknown): RhwpFieldTargetV1 {
  const target = record(value, 'target');
  if (target.kind === 'table_cell_text') return parseTableCellTextTarget(target);
  if (target.kind === 'table_cell_region') return parseTableCellRegionTarget(target);
  if (target.kind === 'form_text') return parseFormTextTarget(target);
  throw new DocumentAgentError(
    'INVALID_COMMAND',
    'target.kind는 table_cell_text, table_cell_region 또는 form_text여야 합니다.',
  );
}

export function parseApplyTextCommand(value: unknown): RhwpApplyTextCommandV1 {
  const command = record(value, 'command');
  exactKeys(command, [
    'schemaVersion', 'commandId', 'expectedDocumentEpoch', 'expectedChangeSeq',
    'expectedDocumentSha256', 'target', 'expectedBeforeSha256',
    'expectedFormatSha256', 'expectedAdjacentContextSha256', 'replacement',
  ], 'command');
  if (command.schemaVersion !== 1) {
    throw new DocumentAgentError('INVALID_COMMAND', 'schemaVersion은 1이어야 합니다.');
  }
  commandId(command.commandId);
  safeInteger(command.expectedDocumentEpoch, 1, 'expectedDocumentEpoch');
  safeInteger(command.expectedChangeSeq, 0, 'expectedChangeSeq');
  digest(command.expectedDocumentSha256, 'expectedDocumentSha256');
  const target = parseBodyParagraphTarget(command.target);
  digest(command.expectedBeforeSha256, 'expectedBeforeSha256');
  digest(command.expectedFormatSha256, 'expectedFormatSha256');
  digest(command.expectedAdjacentContextSha256, 'expectedAdjacentContextSha256');
  if (typeof command.replacement !== 'string'
      || Array.from(command.replacement).length > 4000) {
    throw new DocumentAgentError('INVALID_COMMAND', 'replacement는 4000자 이하 문자열이어야 합니다.');
  }
  if (/[\u0000-\u001f\u007f]/u.test(command.replacement)) {
    throw new DocumentAgentError('INVALID_COMMAND', 'replacement에 control 문자를 넣을 수 없습니다.');
  }
  return { ...command, target } as unknown as RhwpApplyTextCommandV1;
}

export function parseApplyFieldCommand(value: unknown): RhwpApplyFieldCommandV1 {
  const command = record(value, 'command');
  exactKeys(command, [
    'schemaVersion', 'commandId', 'expectedDocumentEpoch', 'expectedChangeSeq',
    'expectedDocumentSha256', 'target', 'expectedBeforeSha256',
    'expectedFormatSha256', 'expectedAdjacentContextSha256', 'replacement',
  ], 'command');
  if (command.schemaVersion !== 1) {
    throw new DocumentAgentError('INVALID_COMMAND', 'schemaVersion은 1이어야 합니다.');
  }
  commandId(command.commandId);
  safeInteger(command.expectedDocumentEpoch, 1, 'expectedDocumentEpoch');
  safeInteger(command.expectedChangeSeq, 0, 'expectedChangeSeq');
  digest(command.expectedDocumentSha256, 'expectedDocumentSha256');
  const target = parseFieldTarget(command.target);
  digest(command.expectedBeforeSha256, 'expectedBeforeSha256');
  digest(command.expectedFormatSha256, 'expectedFormatSha256');
  digest(command.expectedAdjacentContextSha256, 'expectedAdjacentContextSha256');
  if (typeof command.replacement !== 'string'
      || Array.from(command.replacement).length > 4000) {
    throw new DocumentAgentError('INVALID_COMMAND', 'replacement는 4000자 이하 문자열이어야 합니다.');
  }
  if (/[\u0000-\u0009\u000b\u000c\u000e-\u001f\u007f]/u.test(command.replacement)) {
    throw new DocumentAgentError('INVALID_COMMAND', 'replacement에 control 문자를 넣을 수 없습니다.');
  }
  if (/\r/u.test(command.replacement)
      || (target.kind !== 'table_cell_region' && /\n/u.test(command.replacement))) {
    throw new DocumentAgentError('INVALID_COMMAND', 'atomic text field에는 줄바꿈을 넣을 수 없습니다.');
  }
  return { ...command, target } as unknown as RhwpApplyFieldCommandV1;
}

export function parseRevertTextCommand(value: unknown): RhwpRevertTextCommandV1 {
  const command = record(value, 'command');
  exactKeys(command, [
    'schemaVersion', 'commandId', 'expectedDocumentEpoch', 'expectedChangeSeq',
    'expectedAfterDocumentSha256', 'expectedAfterSha256',
  ], 'command');
  if (command.schemaVersion !== 1) {
    throw new DocumentAgentError('INVALID_COMMAND', 'schemaVersion은 1이어야 합니다.');
  }
  commandId(command.commandId);
  safeInteger(command.expectedDocumentEpoch, 1, 'expectedDocumentEpoch');
  safeInteger(command.expectedChangeSeq, 0, 'expectedChangeSeq');
  digest(command.expectedAfterDocumentSha256, 'expectedAfterDocumentSha256');
  digest(command.expectedAfterSha256, 'expectedAfterSha256');
  return command as unknown as RhwpRevertTextCommandV1;
}

export function parseRevertFieldCommand(value: unknown): RhwpRevertFieldCommandV1 {
  return parseRevertTextCommand(value) as RhwpRevertFieldCommandV1;
}

export function assertEmptyParams(value: Record<string, unknown>, label: string): void {
  exactKeys(value, [], label);
}

export function assertOnlyParam(
  value: Record<string, unknown>,
  key: string,
  label: string,
): unknown {
  exactKeys(value, [key], label);
  return value[key];
}
