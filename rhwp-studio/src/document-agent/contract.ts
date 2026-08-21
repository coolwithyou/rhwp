import {
  DocumentAgentError,
  type RhwpApplyFieldCommandV1,
  type RhwpApplyTextCommandV1,
  type RhwpBodyParagraphTargetV1,
  type RhwpFieldTargetV1,
  type RhwpFieldRestoreFormatV1,
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

function restoreCharShapeIds(value: unknown, label: string): number[] {
  if (!Array.isArray(value) || value.length < 1 || value.length > 4000) {
    throw new DocumentAgentError('INVALID_COMMAND', `${label}가 올바르지 않습니다.`);
  }
  for (const [index, id] of value.entries()) safeInteger(id, 0, `${label}[${index}]`);
  return value as number[];
}

function parseFieldRestoreFormat(value: unknown, target: RhwpFieldTargetV1): RhwpFieldRestoreFormatV1 {
  const format = record(value, 'replacementFormat');
  if (format.kind !== target.kind) {
    throw new DocumentAgentError('INVALID_COMMAND', 'replacementFormat kind가 target과 다릅니다.');
  }
  if (format.kind === 'table_cell_region') {
    exactKeys(format, ['kind', 'paragraphs'], 'replacementFormat');
    if (!Array.isArray(format.paragraphs) || format.paragraphs.length < 1 || format.paragraphs.length > 100) {
      throw new DocumentAgentError('INVALID_COMMAND', 'replacementFormat.paragraphs가 올바르지 않습니다.');
    }
    const paragraphs = format.paragraphs.map((value, index) => {
      const paragraph = record(value, `replacementFormat.paragraphs[${index}]`);
      exactKeys(paragraph, ['length', 'charShapeIds', 'paraShapeId'], `replacementFormat.paragraphs[${index}]`);
      safeInteger(paragraph.length, 0, `replacementFormat.paragraphs[${index}].length`);
      safeInteger(paragraph.paraShapeId, 0, `replacementFormat.paragraphs[${index}].paraShapeId`);
      return {
        length: paragraph.length,
        charShapeIds: restoreCharShapeIds(
          paragraph.charShapeIds,
          `replacementFormat.paragraphs[${index}].charShapeIds`,
        ),
        paraShapeId: paragraph.paraShapeId,
      };
    });
    return { kind: 'table_cell_region', paragraphs };
  }
  const allowed = format.kind === 'form_text'
    ? ['kind', 'charShapeIds', 'paraShapeId', 'styleId']
    : ['kind', 'charShapeIds', 'paraShapeId'];
  exactKeys(format, allowed, 'replacementFormat');
  safeInteger(format.paraShapeId, 0, 'replacementFormat.paraShapeId');
  const charShapeIds = restoreCharShapeIds(format.charShapeIds, 'replacementFormat.charShapeIds');
  if (format.kind === 'form_text') {
    safeInteger(format.styleId, 0, 'replacementFormat.styleId');
    return {
      kind: 'form_text',
      charShapeIds,
      paraShapeId: format.paraShapeId,
      styleId: format.styleId,
    };
  }
  return { kind: 'table_cell_text', charShapeIds, paraShapeId: format.paraShapeId };
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
  const allowedKeys = [
    'schemaVersion', 'commandId', 'expectedDocumentEpoch', 'expectedChangeSeq',
    'expectedDocumentSha256', 'target', 'expectedBeforeSha256',
    'expectedFormatSha256', 'expectedAdjacentContextSha256', 'replacement',
  ];
  const optionalKeys = [
    ...(command.replacementStyle === undefined ? [] : ['replacementStyle']),
    ...(command.replacementFormat === undefined ? [] : ['replacementFormat']),
    ...(command.expectedReplacementFormatSha256 === undefined ? [] : ['expectedReplacementFormatSha256']),
  ];
  exactKeys(command, [...allowedKeys, ...optionalKeys], 'command');
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
  if (command.replacementStyle !== undefined
      && command.replacementStyle !== 'actual-input'
      && command.replacementStyle !== 'preserve'
      && command.replacementStyle !== 'restore-exact') {
    throw new DocumentAgentError('INVALID_COMMAND', 'replacementStyle이 올바르지 않습니다.');
  }
  if (command.replacementStyle === 'restore-exact') {
    digest(command.expectedReplacementFormatSha256, 'expectedReplacementFormatSha256');
    const replacementFormat = parseFieldRestoreFormat(command.replacementFormat, target);
    const lengths = target.kind === 'table_cell_region'
      ? command.replacement.split('\n').map(part => Array.from(part).length)
      : [Array.from(command.replacement).length];
    const formats = replacementFormat.kind === 'table_cell_region'
      ? replacementFormat.paragraphs
      : [{ length: lengths[0]!, charShapeIds: replacementFormat.charShapeIds }];
    if (formats.length !== lengths.length || formats.some((format, index) =>
      format.length !== lengths[index]
      || format.charShapeIds.length !== Math.max(lengths[index]!, 1))) {
      throw new DocumentAgentError('INVALID_COMMAND', 'replacementFormat 길이가 replacement와 다릅니다.');
    }
    return { ...command, target, replacementFormat } as unknown as RhwpApplyFieldCommandV1;
  }
  if (command.replacementFormat !== undefined || command.expectedReplacementFormatSha256 !== undefined) {
    throw new DocumentAgentError('INVALID_COMMAND', 'exact 복원 서식은 restore-exact 명령에서만 사용할 수 있습니다.');
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
