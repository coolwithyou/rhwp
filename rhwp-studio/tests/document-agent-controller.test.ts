import test from 'node:test';
import assert from 'node:assert/strict';

import { EventBus } from '../src/core/event-bus.ts';
import {
  DocumentAgentController,
  collectFieldTargetEvidence,
  collectTargetEvidence,
  type DocumentAgentWasm,
  type DocumentAgentInput,
  type RhwpFormFieldEntry,
} from '../src/document-agent/controller.ts';
import type {
  RhwpApplyFieldCommandV1,
  RhwpApplyTextCommandV1,
  RhwpBodyParagraphTargetV1,
} from '../src/document-agent/types.ts';

type Paragraph = {
  text: string;
  paraShapeId: number;
  styleId: number;
  charShapeIds: number[];
  controls: number[];
  fields: Array<[number, number]>;
};

type CellParagraph = Pick<Paragraph, 'text' | 'paraShapeId' | 'charShapeIds'>;

const encoder = new TextEncoder();

class FakeWasm implements DocumentAgentWasm {
  documentGeneration = 1;
  pageCount = 2;
  sourceFormat = 'hwp';
  paragraphs: Paragraph[][] = [[
    paragraph('앞 문단'),
    paragraph('기존 문단'),
    paragraph('뒤 문단'),
  ]];
  cells: CellParagraph[][] = [
    [cellParagraph('항목명')],
    [cellParagraph('기존 값')],
  ];
  formFields: RhwpFormFieldEntry[] = [];
  replacementPageCount: number | null = null;

  getSourceFormat() { return this.sourceFormat; }
  getSectionCount() { return this.paragraphs.length; }
  getParagraphCount(section: number) { return this.paragraphs[section]?.length ?? 0; }
  getParagraphLength(section: number, paragraphIndex: number) {
    return Array.from(this.paragraphs[section][paragraphIndex].text).length;
  }
  getTextRange(section: number, paragraphIndex: number, offset: number, count: number) {
    return Array.from(this.paragraphs[section][paragraphIndex].text)
      .slice(offset, offset + count)
      .join('');
  }
  getControlTextPositions(section: number, paragraphIndex: number) {
    return [...this.paragraphs[section][paragraphIndex].controls];
  }
  getCharPropertiesAt(section: number, paragraphIndex: number, offset: number) {
    const para = this.paragraphs[section][paragraphIndex];
    return { charShapeId: para.charShapeIds[Math.min(offset, para.charShapeIds.length - 1)] ?? 1 };
  }
  getParaPropertiesAt(section: number, paragraphIndex: number) {
    return { paraShapeId: this.paragraphs[section][paragraphIndex].paraShapeId };
  }
  getStyleAt(section: number, paragraphIndex: number) {
    return { id: this.paragraphs[section][paragraphIndex].styleId, name: '본문' };
  }
  getFieldInfoAt(pos: { sectionIndex: number; paragraphIndex: number; charOffset: number }) {
    const exactFormField = this.formFields.find(field =>
      field.location.sectionIndex === pos.sectionIndex
      && field.location.paraIndex === pos.paragraphIndex
      && pos.charOffset >= (field.startCharIdx ?? -1)
      && pos.charOffset <= (field.endCharIdx ?? -1));
    if (exactFormField) {
      return {
        inField: true,
        fieldId: exactFormField.fieldId,
        fieldType: exactFormField.fieldType,
        startCharIdx: exactFormField.startCharIdx,
        endCharIdx: exactFormField.endCharIdx,
        editableInForm: exactFormField.editableInForm,
      };
    }
    const fields = this.paragraphs[pos.sectionIndex][pos.paragraphIndex].fields;
    const field = fields.find(([start, end]) => pos.charOffset >= start && pos.charOffset <= end);
    return field
      ? { inField: true, startCharIdx: field[0], endCharIdx: field[1] }
      : { inField: false };
  }
  getFieldList() { return structuredClone(this.formFields); }
  getFieldValue(fieldId: number) {
    const field = this.formFields.find(entry => entry.fieldId === fieldId);
    return field ? { ok: true, value: field.value } : { ok: false, value: '' };
  }
  setFieldValue(fieldId: number, value: string) {
    const field = this.formFields.find(entry => entry.fieldId === fieldId);
    if (!field || field.startCharIdx === undefined || field.endCharIdx === undefined) {
      return { ok: false, fieldId, oldValue: '', newValue: '' };
    }
    const paragraph = this.paragraphs[field.location.sectionIndex][field.location.paraIndex];
    const oldValue = field.value;
    const chars = Array.from(paragraph.text);
    const shapes = [...paragraph.charShapeIds];
    const replacement = Array.from(value);
    const inheritedShape = shapes[field.startCharIdx] ?? shapes[field.startCharIdx - 1] ?? 4;
    chars.splice(field.startCharIdx, field.endCharIdx - field.startCharIdx, ...replacement);
    shapes.splice(
      field.startCharIdx,
      field.endCharIdx - field.startCharIdx,
      ...Array(Math.max(replacement.length, 1)).fill(inheritedShape),
    );
    paragraph.text = chars.join('');
    paragraph.charShapeIds = shapes.slice(0, Math.max(chars.length, 1));
    field.endCharIdx = field.startCharIdx + replacement.length;
    field.value = value;
    return { ok: true, fieldId, oldValue, newValue: value };
  }
  getTextInCell(_section: number, _parentPara: number, _controlIndex: number, cellIndex: number, cellPara: number, offset: number, count: number) {
    return Array.from(this.cells[cellIndex][cellPara].text).slice(offset, offset + count).join('');
  }
  getCellParagraphLength(_section: number, _parentPara: number, _controlIndex: number, cellIndex: number, cellPara: number) {
    return Array.from(this.cells[cellIndex][cellPara].text).length;
  }
  getCellParagraphCount(_section: number, _parentPara: number, _controlIndex: number, cellIndex: number) {
    return this.cells[cellIndex]?.length ?? 0;
  }
  getCellCharPropertiesAt(_section: number, _parentPara: number, _controlIndex: number, cellIndex: number, cellPara: number, offset: number) {
    const paragraph = this.cells[cellIndex][cellPara];
    return { charShapeId: paragraph.charShapeIds[Math.min(offset, paragraph.charShapeIds.length - 1)] ?? 4 };
  }
  getCellParaPropertiesAt(_section: number, _parentPara: number, _controlIndex: number, cellIndex: number, cellPara: number) {
    return { paraShapeId: this.cells[cellIndex][cellPara].paraShapeId };
  }
  getCellOwnProperties(_section: number, _parentPara: number, _controlIndex: number, cellIndex: number) {
    return { width: 100 + cellIndex, fillColor: '#ffffff' };
  }
  getTableDimensions() { return { rowCount: 1, colCount: 2, cellCount: 2 }; }
  replaceText(section: number, paragraphIndex: number, offset: number, length: number, text: string) {
    const para = this.paragraphs[section][paragraphIndex];
    if (offset !== 0 || length !== Array.from(para.text).length) return { ok: false };
    para.text = text;
    const newLength = Array.from(text).length;
    para.charShapeIds = Array(Math.max(newLength, 1)).fill(para.charShapeIds[0] ?? 1);
    if (this.replacementPageCount !== null) this.pageCount = this.replacementPageCount;
    return { ok: true, charOffset: 0, newLength };
  }
  setCharShapeId(section: number, paragraphIndex: number, start: number, end: number, id: number) {
    const para = this.paragraphs[section][paragraphIndex];
    for (let index = start; index < end; index += 1) para.charShapeIds[index] = id;
    return '{}';
  }
  setParaShapeId(section: number, paragraphIndex: number, id: number) {
    this.paragraphs[section][paragraphIndex].paraShapeId = id;
    return '{}';
  }
  replaceTextInCellDeferredPagination(_section: number, _parentPara: number, _controlIndex: number, cellIndex: number, cellPara: number, offset: number, deleteCount: number, text: string) {
    const paragraph = this.cells[cellIndex][cellPara];
    const chars = Array.from(paragraph.text);
    chars.splice(offset, deleteCount, ...Array.from(text));
    paragraph.text = chars.join('');
    paragraph.charShapeIds = Array(Math.max(chars.length, 1)).fill(paragraph.charShapeIds[0] ?? 4);
    if (this.replacementPageCount !== null) this.pageCount = this.replacementPageCount;
    return { ok: true, charOffset: offset + Array.from(text).length, paginationDeferred: true };
  }
  insertTextInCellDeferredPagination(section: number, parentPara: number, controlIndex: number, cellIndex: number, cellPara: number, offset: number, text: string) {
    return this.replaceTextInCellDeferredPagination(section, parentPara, controlIndex, cellIndex, cellPara, offset, 0, text);
  }
  deleteTextInCellDeferredPagination(section: number, parentPara: number, controlIndex: number, cellIndex: number, cellPara: number, offset: number, count: number) {
    return this.replaceTextInCellDeferredPagination(section, parentPara, controlIndex, cellIndex, cellPara, offset, count, '');
  }
  setCharShapeIdInCell(_section: number, _parentPara: number, _controlIndex: number, cellIndex: number, cellPara: number, start: number, end: number, id: number) {
    const paragraph = this.cells[cellIndex][cellPara];
    for (let index = start; index < end; index += 1) paragraph.charShapeIds[index] = id;
    return '{}';
  }
  setCellParaShapeId(_section: number, _parentPara: number, _controlIndex: number, cellIndex: number, cellPara: number, id: number) {
    this.cells[cellIndex][cellPara].paraShapeId = id;
    return '{}';
  }
  splitParagraphInCell(_section: number, _parentPara: number, _controlIndex: number, cellIndex: number, cellPara: number, charOffset: number) {
    const paragraph = this.cells[cellIndex][cellPara];
    const chars = Array.from(paragraph.text);
    const before = chars.slice(0, charOffset).join('');
    const after = chars.slice(charOffset).join('');
    const beforeShapes = paragraph.charShapeIds.slice(0, Math.max(charOffset, 1));
    const afterShapes = paragraph.charShapeIds.slice(charOffset);
    paragraph.text = before;
    paragraph.charShapeIds = beforeShapes.length > 0 ? beforeShapes : [7];
    this.cells[cellIndex].splice(cellPara + 1, 0, {
      text: after,
      paraShapeId: paragraph.paraShapeId,
      charShapeIds: afterShapes.length > 0 ? afterShapes : [paragraph.charShapeIds[0] ?? 7],
    });
    return '{}';
  }
  mergeParagraphInCell(_section: number, _parentPara: number, _controlIndex: number, cellIndex: number, cellPara: number) {
    const previous = this.cells[cellIndex][cellPara - 1];
    const [removed] = this.cells[cellIndex].splice(cellPara, 1);
    previous.text += removed.text;
    previous.charShapeIds = [
      ...previous.charShapeIds.slice(0, Math.max(Array.from(previous.text).length - Array.from(removed.text).length, 0)),
      ...removed.charShapeIds.slice(0, Math.max(Array.from(removed.text).length, 1)),
    ];
    return JSON.stringify({ removedParaMeta: {} });
  }
  beginDeferredPagination() {}
  flushDeferredPagination() {}
  cancelDeferredPagination() {}
  getPageOfPosition() { return { ok: true, page: 1 }; }
  exportHwp() { return this.borrowDocumentHandle().exportHwp(); }
  exportHwpx() { return this.borrowDocumentHandle().exportHwpx(); }
  borrowDocumentHandle() {
    return {
      exportHwp: () => encoder.encode(JSON.stringify({
        paragraphs: this.paragraphs,
        cells: this.cells,
        formFields: this.formFields,
        pageCount: this.pageCount,
      })),
      exportHwpx: () => encoder.encode(JSON.stringify({
        paragraphs: this.paragraphs,
        cells: this.cells,
        formFields: this.formFields,
        pageCount: this.pageCount,
      })),
    };
  }

  cloneState() {
    return {
      pageCount: this.pageCount,
      paragraphs: structuredClone(this.paragraphs),
      cells: structuredClone(this.cells),
      formFields: structuredClone(this.formFields),
    };
  }
  restoreState(snapshot: ReturnType<FakeWasm['cloneState']>) {
    this.pageCount = snapshot.pageCount;
    this.paragraphs = structuredClone(snapshot.paragraphs);
    this.cells = structuredClone(snapshot.cells);
    this.formFields = structuredClone(snapshot.formFields);
  }
}

class FakeInput implements DocumentAgentInput {
  transactions = 0;
  position = { sectionIndex: 0, paragraphIndex: 1, charOffset: 0 };
  selection: { start: typeof this.position; end: typeof this.position } | null = null;
  private wasm: FakeWasm;
  private eventBus: EventBus;

  constructor(wasm: FakeWasm, eventBus: EventBus) {
    this.wasm = wasm;
    this.eventBus = eventBus;
  }

  getCursorPosition() { return { ...this.position }; }
  getSelection() { return this.selection ? structuredClone(this.selection) : null; }
  async executeDocumentAgentOperation(
    desc: Parameters<DocumentAgentInput['executeDocumentAgentOperation']>[0],
    render: () => Promise<void>,
  ) {
    const snapshot = this.wasm.cloneState();
    const positionBefore = { ...this.position };
    try {
      const result = desc.operation(this.wasm);
      if (result === null) return;
      this.position = { ...result };
      try {
        await render();
      } catch (cause) {
        this.wasm.restoreState(snapshot);
        this.position = positionBefore;
        throw Object.assign(new Error('render failed', { cause }), {
          code: 'RENDER_FAILED',
          recovered: true,
        });
      }
      this.transactions += 1;
      this.eventBus.emit('document-mutated', desc.operationType);
    } catch (error) {
      this.wasm.restoreState(snapshot);
      this.position = positionBefore;
      throw error;
    }
  }
  focusBodyParagraph(section: number, paragraphIndex: number, length: number) {
    this.position = { sectionIndex: section, paragraphIndex, charOffset: length };
    this.selection = {
      start: { sectionIndex: section, paragraphIndex, charOffset: 0 },
      end: { sectionIndex: section, paragraphIndex, charOffset: length },
    };
    return true;
  }
  focusTableCellText(
    section: number,
    parentPara: number,
    controlIndex: number,
    cellIndex: number,
    cellParagraph: number,
  ) {
    this.position = {
      sectionIndex: section,
      paragraphIndex: cellParagraph,
      charOffset: 0,
      parentParaIndex: parentPara,
      controlIndex,
      cellIndex,
      cellParaIndex: cellParagraph,
    };
    return { focused: true, page: 3 };
  }
  focusFormText(section: number, paragraphIndex: number, fieldId: number) {
    const field = this.wasm.formFields.find(entry => entry.fieldId === fieldId)!;
    this.position = {
      sectionIndex: section,
      paragraphIndex,
      charOffset: field.endCharIdx ?? 0,
    };
    this.selection = {
      start: { sectionIndex: section, paragraphIndex, charOffset: field.startCharIdx ?? 0 },
      end: { sectionIndex: section, paragraphIndex, charOffset: field.endCharIdx ?? 0 },
    };
    return { focused: true, page: 2 };
  }
}

function paragraph(text: string): Paragraph {
  return {
    text,
    paraShapeId: 2,
    styleId: 3,
    charShapeIds: Array(Math.max(Array.from(text).length, 1)).fill(4),
    controls: [],
    fields: [],
  };
}

function cellParagraph(text: string): CellParagraph {
  return {
    text,
    paraShapeId: 6,
    charShapeIds: Array(Math.max(Array.from(text).length, 1)).fill(7),
  };
}

function target(wasm: FakeWasm): RhwpBodyParagraphTargetV1 {
  return {
    kind: 'body_paragraph',
    section: 0,
    paragraph: 1,
    charOffset: 0,
    length: wasm.getParagraphLength(0, 1),
  };
}

function harness(options: { render?: () => Promise<void> } = {}) {
  const wasm = new FakeWasm();
  const eventBus = new EventBus();
  const input = new FakeInput(wasm, eventBus);
  const events: unknown[] = [];
  eventBus.on('document-agent-changed', (event) => events.push(event));
  const controller = new DocumentAgentController({
    wasm,
    input,
    eventBus,
    isDirty: () => true,
    render: options.render ?? (async () => {}),
  });
  return { controller, wasm, input, events, eventBus };
}

function applyCommand(
  controller: DocumentAgentController,
  wasm: FakeWasm,
  replacement = '새 문단',
): RhwpApplyTextCommandV1 {
  const state = controller.getDocumentState();
  const exactTarget = target(wasm);
  const evidence = collectTargetEvidence(wasm, exactTarget);
  return {
    schemaVersion: 1,
    commandId: 'cmd-1',
    expectedDocumentEpoch: state.documentEpoch,
    expectedChangeSeq: state.changeSeq,
    expectedDocumentSha256: state.documentSha256,
    target: exactTarget,
    expectedBeforeSha256: evidence.textSha256,
    expectedFormatSha256: evidence.formatSha256,
    expectedAdjacentContextSha256: evidence.adjacentContextSha256,
    replacement,
  };
}

function fieldCommand(
  controller: DocumentAgentController,
  wasm: FakeWasm,
  replacement = '새 필드 값',
): RhwpApplyFieldCommandV1 {
  const state = controller.getDocumentState();
  const exactTarget = {
    kind: 'table_cell_text' as const,
    section: 0,
    parentPara: 1,
    controlIndex: 0,
    cellIndex: 1,
    cellParagraph: 0,
  };
  const evidence = collectFieldTargetEvidence(wasm, exactTarget);
  return {
    schemaVersion: 1,
    commandId: 'field-cmd-1',
    expectedDocumentEpoch: state.documentEpoch,
    expectedChangeSeq: state.changeSeq,
    expectedDocumentSha256: state.documentSha256,
    target: exactTarget,
    expectedBeforeSha256: evidence.textSha256,
    expectedFormatSha256: evidence.formatSha256,
    expectedAdjacentContextSha256: evidence.adjacentContextSha256,
    replacement,
  };
}

function fieldRegionCommand(
  controller: DocumentAgentController,
  wasm: FakeWasm,
  replacement = '첫 문단\n둘째 문단',
): RhwpApplyFieldCommandV1 {
  const state = controller.getDocumentState();
  const exactTarget = {
    kind: 'table_cell_region' as const,
    section: 0,
    parentPara: 1,
    controlIndex: 0,
    cellIndex: 1,
  };
  const evidence = collectFieldTargetEvidence(wasm, exactTarget);
  return {
    schemaVersion: 1,
    commandId: 'field-region-cmd-1',
    expectedDocumentEpoch: state.documentEpoch,
    expectedChangeSeq: state.changeSeq,
    expectedDocumentSha256: state.documentSha256,
    target: exactTarget,
    expectedBeforeSha256: evidence.textSha256,
    expectedFormatSha256: evidence.formatSha256,
    expectedAdjacentContextSha256: evidence.adjacentContextSha256,
    replacement,
  };
}

function installFormTextField(wasm: FakeWasm, value = '기존 회사명') {
  const prefix = '회사명: ';
  wasm.paragraphs[0][1] = paragraph(`${prefix}${value} / 확인`);
  wasm.formFields = [{
    fieldId: 41,
    fieldType: 'clickhere',
    cellField: false,
    name: '회사명',
    guide: '회사명을 입력하세요',
    command: 'field:company-name',
    value,
    location: { sectionIndex: 0, paraIndex: 1 },
    startCharIdx: Array.from(prefix).length,
    endCharIdx: Array.from(prefix).length + Array.from(value).length,
    editableInForm: true,
  }];
}

function formFieldCommand(
  controller: DocumentAgentController,
  wasm: FakeWasm,
  replacement = '주식회사 노튼',
): RhwpApplyFieldCommandV1 {
  const state = controller.getDocumentState();
  const exactTarget = {
    kind: 'form_text' as const,
    section: 0,
    paragraph: 1,
    fieldId: 41,
  };
  const evidence = collectFieldTargetEvidence(wasm, exactTarget);
  return {
    schemaVersion: 1,
    commandId: 'form-field-cmd-1',
    expectedDocumentEpoch: state.documentEpoch,
    expectedChangeSeq: state.changeSeq,
    expectedDocumentSha256: state.documentSha256,
    target: exactTarget,
    expectedBeforeSha256: evidence.textSha256,
    expectedFormatSha256: evidence.formatSha256,
    expectedAdjacentContextSha256: evidence.adjacentContextSha256,
    replacement,
  };
}

test('field apply/revert는 exact 셀만 한 트랜잭션으로 변경하고 복원한다', async () => {
  const { controller, wasm, input, events } = harness();
  const replacement = 'AI 기반 사업계획서 작성 서비스';
  const command = fieldCommand(controller, wasm, replacement);
  const labelBefore = structuredClone(wasm.cells[0]);

  const applied = await controller.applyFieldCommand(command);
  assert.equal(wasm.cells[1][0].text, replacement);
  assert.deepEqual(wasm.cells[0], labelBefore);
  assert.equal(input.transactions, 1);
  assert.equal(applied.afterChangeSeq, 1);
  assert.equal((events[0] as { reason: string }).reason, 'field_agent_apply');

  const reverted = await controller.revertFieldCommand({
    schemaVersion: 1,
    commandId: command.commandId,
    expectedDocumentEpoch: applied.documentEpoch,
    expectedChangeSeq: applied.afterChangeSeq,
    expectedAfterDocumentSha256: applied.afterDocumentSha256,
    expectedAfterSha256: applied.afterTextSha256,
  });
  assert.equal(wasm.cells[1][0].text, '기존 값');
  assert.deepEqual(wasm.cells[0], labelBefore);
  assert.equal(input.transactions, 2);
  assert.equal(reverted.afterChangeSeq, 2);
  assert.equal((events[1] as { reason: string }).reason, 'field_agent_revert');
});

test('table_cell_region apply/revert는 셀 전체 장문 문단을 변경하고 원문 구조를 복원한다', async () => {
  const { controller, wasm, input } = harness();
  wasm.cells[1] = [cellParagraph('기존 첫 문단'), cellParagraph('기존 둘째 문단')];
  const before = structuredClone(wasm.cells);
  const command = fieldRegionCommand(controller, wasm, '창업 경험을 바탕으로 문제를 발견했습니다.\n팀의 전문성으로 해결하겠습니다.');

  const applied = await controller.applyFieldCommand(command);
  assert.deepEqual(wasm.cells[1].map(paragraph => paragraph.text), [
    '창업 경험을 바탕으로 문제를 발견했습니다.',
    '팀의 전문성으로 해결하겠습니다.',
  ]);
  assert.deepEqual(wasm.cells[0], before[0]);
  assert.equal(applied.target.kind, 'table_cell_region');
  assert.equal(input.transactions, 1);

  const reverted = await controller.revertFieldCommand({
    schemaVersion: 1,
    commandId: command.commandId,
    expectedDocumentEpoch: applied.documentEpoch,
    expectedChangeSeq: applied.afterChangeSeq,
    expectedAfterDocumentSha256: applied.afterDocumentSha256,
    expectedAfterSha256: applied.afterTextSha256,
  });
  assert.deepEqual(wasm.cells, before);
  assert.equal(reverted.target.kind, 'table_cell_region');
  assert.equal(input.transactions, 2);
});

test('form_text apply/revert는 exact 누름틀 값만 변경하고 구조와 문맥을 복원한다', async () => {
  const { controller, wasm, input, events } = harness();
  installFormTextField(wasm);
  const beforeParagraphs = structuredClone(wasm.paragraphs);
  const command = formFieldCommand(controller, wasm);
  const beforeEvidence = collectFieldTargetEvidence(wasm, command.target);

  const applied = await controller.applyFieldCommand(command);
  assert.equal(wasm.getFieldValue(41).value, '주식회사 노튼');
  assert.equal(wasm.paragraphs[0][1].text, '회사명: 주식회사 노튼 / 확인');
  assert.equal(applied.adjacentContextSha256, beforeEvidence.adjacentContextSha256);
  assert.equal(applied.formatSha256, beforeEvidence.formatSha256);
  assert.equal(input.transactions, 1);

  const reverted = await controller.revertFieldCommand({
    schemaVersion: 1,
    commandId: command.commandId,
    expectedDocumentEpoch: applied.documentEpoch,
    expectedChangeSeq: applied.afterChangeSeq,
    expectedAfterDocumentSha256: applied.afterDocumentSha256,
    expectedAfterSha256: applied.afterTextSha256,
  });
  assert.equal(wasm.getFieldValue(41).value, '기존 회사명');
  assert.deepEqual(wasm.paragraphs, beforeParagraphs);
  assert.equal(reverted.adjacentContextSha256, beforeEvidence.adjacentContextSha256);
  assert.equal(reverted.formatSha256, beforeEvidence.formatSha256);
  assert.equal(input.transactions, 2);
  assert.deepEqual(events.map(event => (event as { reason: string }).reason), [
    'field_agent_apply',
    'field_agent_revert',
  ]);
});

test('선택 마커 같은 길이 치환은 혼합 글자 서식을 글자별로 보존하고 revert한다', async () => {
  const { controller, wasm, input } = harness();
  wasm.cells[1][0] = cellParagraph('□ 예비창업자 □ 폐업 후 재창업자');
  const originalShapeIds = Array.from(wasm.cells[1][0].text, (_, index) => index < 8 ? 7 : 9);
  wasm.cells[1][0].charShapeIds = [...originalShapeIds];
  const command = fieldCommand(controller, wasm, '■ 예비창업자 □ 폐업 후 재창업자');

  const applied = await controller.applyFieldCommand(command);
  assert.equal(wasm.cells[1][0].text, '■ 예비창업자 □ 폐업 후 재창업자');
  assert.deepEqual(wasm.cells[1][0].charShapeIds, originalShapeIds);

  await controller.revertFieldCommand({
    schemaVersion: 1,
    commandId: command.commandId,
    expectedDocumentEpoch: applied.documentEpoch,
    expectedChangeSeq: applied.afterChangeSeq,
    expectedAfterDocumentSha256: applied.afterDocumentSha256,
    expectedAfterSha256: applied.afterTextSha256,
  });
  assert.equal(wasm.cells[1][0].text, '□ 예비창업자 □ 폐업 후 재창업자');
  assert.deepEqual(wasm.cells[1][0].charShapeIds, originalShapeIds);
  assert.equal(input.transactions, 2);
});

test('혼합 글자 서식 셀의 길이 변경은 mutation 전에 거부한다', async () => {
  const { controller, wasm, input } = harness();
  wasm.cells[1][0].charShapeIds[1] = 99;
  const command = fieldCommand(controller, wasm, '길이가 다른 선택값');

  await assert.rejects(
    controller.applyFieldCommand(command),
    (error: unknown) => (error as { code?: string }).code === 'TARGET_FORMAT_MISMATCH',
  );
  assert.equal(input.transactions, 0);
  assert.equal(wasm.cells[1][0].text, '기존 값');
});

test('apply/revert는 각각 한 트랜잭션·changeSeq 1회·strict receipt로 종결된다', async () => {
  const { controller, wasm, input, events } = harness();
  const command = applyCommand(controller, wasm);

  const applied = await controller.applyTextCommand(command);
  assert.equal(wasm.paragraphs[0][1].text, '새 문단');
  assert.equal(input.transactions, 1);
  assert.equal(applied.beforeChangeSeq, 0);
  assert.equal(applied.afterChangeSeq, 1);
  assert.equal(events.length, 1);

  const reverted = await controller.revertTextCommand({
    schemaVersion: 1,
    commandId: command.commandId,
    expectedDocumentEpoch: applied.documentEpoch,
    expectedChangeSeq: applied.afterChangeSeq,
    expectedAfterDocumentSha256: applied.afterDocumentSha256,
    expectedAfterSha256: applied.afterTextSha256,
  });
  assert.equal(wasm.paragraphs[0][1].text, '기존 문단');
  assert.equal(input.transactions, 2);
  assert.equal(reverted.beforeChangeSeq, 1);
  assert.equal(reverted.afterChangeSeq, 2);
  assert.equal(events.length, 2);
});

test('target evidence는 세션 stable id 없이 고정된 UTF-8 SHA-256 벡터를 사용한다', () => {
  const { wasm } = harness();
  const evidence = collectTargetEvidence(wasm, target(wasm));

  assert.equal(
    evidence.textSha256,
    '73b107c1b8b366ea082beba6cf29568890d48845fa212849f74b516209a6378e',
  );
  assert.equal(
    evidence.formatSha256,
    '5479c59c8c99dde1c9bb45c70a4689eab28233b28dcf31976db9ab8e4b74ec71',
  );
  assert.equal(
    evidence.adjacentContextSha256,
    '17026195f4004b9995060c1dd27a14fb8619c9afc502a371cbdebf94025b7568',
  );
});

test('replacement 길이는 UTF-16 code unit가 아닌 Unicode code point로 계산한다', async () => {
  const { controller, wasm } = harness();
  const applied = await controller.applyTextCommand(applyCommand(controller, wasm, '🚀 새 문단'));

  assert.equal(applied.afterTextSha256.length, 64);
  assert.equal(wasm.getParagraphLength(0, 1), 6);
});

test('stale preimage는 mutation 0회로 거부된다', async () => {
  const { controller, wasm, input, events } = harness();
  const command = applyCommand(controller, wasm);
  command.expectedBeforeSha256 = 'f'.repeat(64);

  await assert.rejects(
    controller.applyTextCommand(command),
    (error: unknown) => (error as { code?: string }).code === 'TARGET_PREIMAGE_MISMATCH',
  );
  assert.equal(input.transactions, 0);
  assert.equal(events.length, 0);
  assert.equal(wasm.paragraphs[0][1].text, '기존 문단');
});

test('mixed format·control·field target은 mutation 전에 거부된다', () => {
  for (const mutate of [
    (wasm: FakeWasm) => { wasm.paragraphs[0][1].charShapeIds[1] = 99; },
    (wasm: FakeWasm) => { wasm.paragraphs[0][1].controls = [1]; },
    (wasm: FakeWasm) => { wasm.paragraphs[0][1].fields = [[0, 2]]; },
  ]) {
    const { controller, wasm, input } = harness();
    mutate(wasm);
    assert.throws(
      () => collectTargetEvidence(wasm, target(wasm)),
      (error: unknown) => (error as { code?: string }).code === 'TARGET_FORMAT_MISMATCH',
    );
    assert.equal(input.transactions, 0);
  }
});

test('page count postcondition 실패는 같은 트랜잭션에서 rollback된다', async () => {
  const { controller, wasm, input, events } = harness();
  const command = applyCommand(controller, wasm, '페이지 증가');
  wasm.replacementPageCount = 3;

  await assert.rejects(
    controller.applyTextCommand(command),
    (error: unknown) => (error as { code?: string }).code === 'PAGE_COUNT_CHANGED',
  );
  assert.equal(wasm.paragraphs[0][1].text, '기존 문단');
  assert.equal(wasm.pageCount, 2);
  assert.equal(input.transactions, 0);
  assert.equal(events.length, 0);
});

test('apply 뒤 일반 편집이 있으면 agent revert를 거부한다', async () => {
  const { controller, wasm, eventBus } = harness();
  const applied = await controller.applyTextCommand(applyCommand(controller, wasm));
  wasm.paragraphs[0][0].text = '사용자 편집';
  eventBus.emit('document-mutated', 'user-edit');

  await assert.rejects(
    controller.revertTextCommand({
      schemaVersion: 1,
      commandId: applied.commandId,
      expectedDocumentEpoch: applied.documentEpoch,
      expectedChangeSeq: applied.afterChangeSeq,
      expectedAfterDocumentSha256: applied.afterDocumentSha256,
      expectedAfterSha256: applied.afterTextSha256,
    }),
    (error: unknown) => (error as { code?: string }).code === 'COMMAND_NOT_LATEST',
  );
});

test('같은 apply replay는 exact terminal 상태에서만 receipt를 재사용한다', async () => {
  const { controller, wasm, input, eventBus } = harness();
  const command = applyCommand(controller, wasm);
  const first = await controller.applyTextCommand(command);
  assert.deepEqual(await controller.applyTextCommand(structuredClone(command)), first);
  assert.equal(input.transactions, 1);

  await assert.rejects(
    controller.applyTextCommand({ ...command, replacement: '다른 문단' }),
    (error: unknown) => (error as { code?: string }).code === 'COMMAND_REPLAY_MISMATCH',
  );

  wasm.paragraphs[0][0].text = '후속 사용자 편집';
  eventBus.emit('document-mutated', 'user-edit');
  await assert.rejects(
    controller.applyTextCommand(structuredClone(command)),
    (error: unknown) => (error as { code?: string }).code === 'COMMAND_REPLAY_MISMATCH',
  );
});

test('같은 revert replay도 exact terminal 상태가 바뀌면 거부한다', async () => {
  const { controller, wasm, eventBus } = harness();
  const applied = await controller.applyTextCommand(applyCommand(controller, wasm));
  const command = {
    schemaVersion: 1 as const,
    commandId: applied.commandId,
    expectedDocumentEpoch: applied.documentEpoch,
    expectedChangeSeq: applied.afterChangeSeq,
    expectedAfterDocumentSha256: applied.afterDocumentSha256,
    expectedAfterSha256: applied.afterTextSha256,
  };
  const reverted = await controller.revertTextCommand(command);
  assert.deepEqual(await controller.revertTextCommand(structuredClone(command)), reverted);

  wasm.paragraphs[0][0].text = 'revert 뒤 사용자 편집';
  eventBus.emit('document-mutated', 'user-edit');
  await assert.rejects(
    controller.revertTextCommand(structuredClone(command)),
    (error: unknown) => (error as { code?: string }).code === 'COMMAND_REPLAY_MISMATCH',
  );
});

test('strict render 실패는 snapshot을 복구하고 recovered=true를 전달한다', async () => {
  const { controller, wasm, input, events } = harness({
    render: async () => { throw new Error('renderer unavailable'); },
  });
  const command = applyCommand(controller, wasm);

  await assert.rejects(
    controller.applyTextCommand(command),
    (error: unknown) => {
      const typed = error as { code?: string; recovered?: boolean };
      return typed.code === 'RENDER_FAILED' && typed.recovered === true;
    },
  );
  assert.equal(wasm.paragraphs[0][1].text, '기존 문단');
  assert.equal(input.transactions, 0);
  assert.equal(events.length, 0);
});

test('selection context와 focus는 body paragraph만 exact하게 노출한다', () => {
  const { controller, wasm, input } = harness();
  const collapsed = controller.getSelectionContext();
  assert.equal(collapsed.collapsed, true);
  assert.deepEqual(collapsed.target, target(wasm));
  assert.equal(collapsed.page, 2);

  assert.deepEqual(controller.focusTarget(target(wasm)), { focused: true, page: 2 });
  const selected = controller.getSelectionContext();
  assert.equal(selected.collapsed, false);
  assert.equal(selected.selectedTextSha256?.length, 64);
  assert.deepEqual(input.selection?.start, {
    sectionIndex: 0,
    paragraphIndex: 1,
    charOffset: 0,
  });
});

test('field target focus는 exact 표 셀 좌표로 이동하고 mutation을 만들지 않는다', () => {
  const { controller, input, events } = harness();
  const beforeTransactions = input.transactions;
  const fieldTarget = {
    kind: 'table_cell_text' as const,
    section: 0,
    parentPara: 5,
    controlIndex: 1,
    cellIndex: 3,
    cellParagraph: 0,
  };
  assert.deepEqual(controller.focusFieldTarget(fieldTarget), { focused: true, page: 3 });
  assert.deepEqual(input.position, {
    sectionIndex: 0,
    paragraphIndex: 0,
    charOffset: 0,
    parentParaIndex: 5,
    controlIndex: 1,
    cellIndex: 3,
    cellParaIndex: 0,
  });
  assert.equal(input.transactions, beforeTransactions);
  assert.equal(events.length, 0);
});

test('field selection context는 현재 표 셀을 exact target으로 노출하고 본문에서는 null이다', () => {
  const { controller, input } = harness();
  assert.deepEqual(controller.getFieldSelectionContext(), {
    schemaVersion: 1,
    documentEpoch: 1,
    changeSeq: 0,
    page: 2,
    editable: false,
    target: null,
  });

  input.position = {
    sectionIndex: 0,
    paragraphIndex: 0,
    charOffset: 0,
    parentParaIndex: 1,
    controlIndex: 0,
    cellIndex: 1,
    cellParaIndex: 0,
  };
  assert.deepEqual(controller.getFieldSelectionContext(), {
    schemaVersion: 1,
    documentEpoch: 1,
    changeSeq: 0,
    page: 2,
    editable: true,
    target: {
      kind: 'table_cell_text',
      section: 0,
      parentPara: 1,
      controlIndex: 0,
      cellIndex: 1,
      cellParagraph: 0,
    },
  });
});

test('field selection과 focus는 본문 누름틀의 exact fieldId를 노출하고 mutation을 만들지 않는다', () => {
  const { controller, wasm, input, events } = harness();
  installFormTextField(wasm);
  input.position = { sectionIndex: 0, paragraphIndex: 1, charOffset: 7 };

  assert.deepEqual(controller.getFieldSelectionContext(), {
    schemaVersion: 1,
    documentEpoch: 1,
    changeSeq: 0,
    page: 2,
    editable: true,
    target: { kind: 'form_text', section: 0, paragraph: 1, fieldId: 41 },
  });
  assert.deepEqual(
    controller.focusFieldTarget({ kind: 'form_text', section: 0, paragraph: 1, fieldId: 41 }),
    { focused: true, page: 2 },
  );
  assert.deepEqual(input.selection, {
    start: { sectionIndex: 0, paragraphIndex: 1, charOffset: 5 },
    end: { sectionIndex: 0, paragraphIndex: 1, charOffset: 11 },
  });
  assert.equal(input.transactions, 0);
  assert.equal(events.length, 0);
});
