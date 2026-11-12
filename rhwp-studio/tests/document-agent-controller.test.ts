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
  charShapeProperties = new Map<number, { textColor: string; italic: boolean }>([
    [1, { textColor: '#000000', italic: false }],
    [4, { textColor: '#000000', italic: false }],
    [7, { textColor: '#000000', italic: false }],
    [9, { textColor: '#000000', italic: false }],
    [37, { textColor: '#0000ff', italic: true }],
    [40, { textColor: '#000000', italic: false }],
  ]);
  nextCharShapeId = 100;

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
    const charShapeId = para.charShapeIds[Math.min(offset, para.charShapeIds.length - 1)] ?? 1;
    return { charShapeId, ...this.charShapeProperties.get(charShapeId) };
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
    const charShapeId = paragraph.charShapeIds[Math.min(offset, paragraph.charShapeIds.length - 1)] ?? 4;
    return { charShapeId, ...this.charShapeProperties.get(charShapeId) };
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
  applyCharFormat(section: number, paragraphIndex: number, start: number, end: number, propsJson: string) {
    const para = this.paragraphs[section][paragraphIndex];
    const patch = JSON.parse(propsJson) as Partial<{ textColor: string; italic: boolean }>;
    for (let index = start; index < end; index += 1) {
      const beforeId = para.charShapeIds[index] ?? 1;
      const nextId = this.nextCharShapeId++;
      this.charShapeProperties.set(nextId, {
        textColor: '#000000',
        italic: false,
        ...this.charShapeProperties.get(beforeId),
        ...patch,
      });
      para.charShapeIds[index] = nextId;
    }
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
  applyCharFormatInCell(_section: number, _parentPara: number, _controlIndex: number, cellIndex: number, cellPara: number, start: number, end: number, propsJson: string) {
    const paragraph = this.cells[cellIndex][cellPara];
    const patch = JSON.parse(propsJson) as Partial<{ textColor: string; italic: boolean }>;
    for (let index = start; index < end; index += 1) {
      const beforeId = paragraph.charShapeIds[index] ?? 4;
      const nextId = this.nextCharShapeId++;
      this.charShapeProperties.set(nextId, {
        textColor: '#000000',
        italic: false,
        ...this.charShapeProperties.get(beforeId),
        ...patch,
      });
      paragraph.charShapeIds[index] = nextId;
    }
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
        charShapeProperties: [...this.charShapeProperties],
        pageCount: this.pageCount,
      })),
      exportHwpx: () => encoder.encode(JSON.stringify({
        paragraphs: this.paragraphs,
        cells: this.cells,
        formFields: this.formFields,
        charShapeProperties: [...this.charShapeProperties],
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
      charShapeProperties: structuredClone(this.charShapeProperties),
      nextCharShapeId: this.nextCharShapeId,
    };
  }
  restoreState(snapshot: ReturnType<FakeWasm['cloneState']>) {
    this.pageCount = snapshot.pageCount;
    this.paragraphs = structuredClone(snapshot.paragraphs);
    this.cells = structuredClone(snapshot.cells);
    this.formFields = structuredClone(snapshot.formFields);
    this.charShapeProperties = structuredClone(snapshot.charShapeProperties);
    this.nextCharShapeId = snapshot.nextCharShapeId;
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
    replacementStyle: 'actual-input',
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
    replacementStyle: 'actual-input',
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
    replacementStyle: 'actual-input',
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

test('table_cell_region은 안내문을 제거하고 AI 본문을 검은색 실제 입력 서식으로 적용한 뒤 exact 복원한다', async () => {
  const { controller, wasm, input } = harness();
  const first = cellParagraph('※ 안내 본문');
  first.charShapeIds = [37, 37, ...Array(Array.from(first.text).length - 2).fill(40)];
  first.paraShapeId = 0;
  const second = cellParagraph('※ 둘째 안내');
  second.charShapeIds = Array(Array.from(second.text).length).fill(37);
  second.paraShapeId = 35;
  wasm.cells[1] = [first, second];
  const before = structuredClone(wasm.cells);
  const command = fieldRegionCommand(controller, wasm, '사용자 사실을 바탕으로 작성한 첫 문단\n검증 가능한 두 번째 문단');
  const beforeEvidence = collectFieldTargetEvidence(wasm, command.target);

  const applied = await controller.applyFieldCommand(command);
  assert.deepEqual(wasm.cells[1].map(paragraph => paragraph.text), [
    '사용자 사실을 바탕으로 작성한 첫 문단',
    '검증 가능한 두 번째 문단',
  ]);
  assert.ok(!wasm.cells[1].some(paragraph => paragraph.text.includes('안내')));
  for (let paragraph = 0; paragraph < wasm.cells[1].length; paragraph += 1) {
    for (let offset = 0; offset < Array.from(wasm.cells[1][paragraph].text).length; offset += 1) {
      const properties = wasm.getCellCharPropertiesAt(0, 1, 0, 1, paragraph, offset);
      assert.equal(properties.textColor, '#000000');
      assert.equal(properties.italic, false);
    }
  }
  assert.notEqual(applied.formatSha256, beforeEvidence.formatSha256);
  assert.equal(input.transactions, 1);

  await controller.revertFieldCommand({
    schemaVersion: 1,
    commandId: command.commandId,
    expectedDocumentEpoch: applied.documentEpoch,
    expectedChangeSeq: applied.afterChangeSeq,
    expectedAfterDocumentSha256: applied.afterDocumentSha256,
    expectedAfterSha256: applied.afterTextSha256,
  });
  assert.deepEqual(wasm.cells, before);
  assert.equal(input.transactions, 2);
});

test('새 Studio 세션도 봉인된 안내문 서식을 restore-exact로 복원한다', async () => {
  const { controller, wasm } = harness();
  const guide = cellParagraph('※ 파란색 안내문');
  guide.charShapeIds = Array(Array.from(guide.text).length).fill(37);
  guide.paraShapeId = 35;
  wasm.cells[1] = [guide];
  const before = structuredClone(wasm.cells[1]);
  const apply = fieldRegionCommand(controller, wasm, '검은색으로 적용할 AI 본문');
  const beforeEvidence = collectFieldTargetEvidence(wasm, apply.target);

  await controller.applyFieldCommand(apply);
  assert.equal(wasm.getCellCharPropertiesAt(0, 1, 0, 1, 0, 0).textColor, '#000000');

  const reloadedEventBus = new EventBus();
  const reloadedInput = new FakeInput(wasm, reloadedEventBus);
  const reloaded = new DocumentAgentController({
    wasm,
    input: reloadedInput,
    eventBus: reloadedEventBus,
    isDirty: () => true,
    render: async () => {},
  });
  const state = reloaded.getDocumentState();
  const currentEvidence = collectFieldTargetEvidence(wasm, apply.target);
  await reloaded.applyFieldCommand({
    schemaVersion: 1,
    commandId: 'field-region-cmd-1:undo',
    expectedDocumentEpoch: state.documentEpoch,
    expectedChangeSeq: state.changeSeq,
    expectedDocumentSha256: state.documentSha256,
    target: apply.target,
    expectedBeforeSha256: currentEvidence.textSha256,
    expectedFormatSha256: currentEvidence.formatSha256,
    expectedAdjacentContextSha256: currentEvidence.adjacentContextSha256,
    replacement: before.map(paragraph => paragraph.text).join('\n'),
    replacementStyle: 'restore-exact',
    replacementFormat: {
      kind: 'table_cell_region',
      paragraphs: before.map(paragraph => ({
        length: Array.from(paragraph.text).length,
        charShapeIds: [...paragraph.charShapeIds],
        paraShapeId: paragraph.paraShapeId,
      })),
    },
    expectedReplacementFormatSha256: beforeEvidence.formatSha256,
  });
  assert.deepEqual(wasm.cells[1], before);
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

// ── 5차: 같은 세션 field revert는 apply snapshot 명령의 before 스냅숏으로 정확히 복원한다 ──

/**
 * core 스냅숏과 "영역 교체가 되살리지 못하는 문단 끝 빈 run"을 흉내 내는 fake.
 * trailingRun은 export bytes에만 나타나고 target evidence(text/format/context)에는 없다.
 * 셀 텍스트를 교체하면 원래 값('11pt')을 잃는다 — 실제 core에서 끝 빈 run이 바뀌는 현상과 같은 모양.
 */
class SnapshotWasm extends FakeWasm {
  trailingRun = '11pt';
  snapshots = new Map<number, ReturnType<SnapshotWasm['cloneState']>>();
  nextSnapshotId = 0;
  restoreCalls: number[] = [];
  cellTextWrites = 0;

  override cloneState() {
    return { ...super.cloneState(), trailingRun: this.trailingRun };
  }
  override restoreState(snapshot: ReturnType<FakeWasm['cloneState']> & { trailingRun?: string }) {
    super.restoreState(snapshot);
    if (snapshot.trailingRun !== undefined) this.trailingRun = snapshot.trailingRun;
  }
  saveSnapshot() {
    const id = this.nextSnapshotId++;
    this.snapshots.set(id, this.cloneState());
    return id;
  }
  restoreSnapshot(id: number) {
    const snapshot = this.snapshots.get(id);
    if (!snapshot) throw new Error(`snapshot ${id} 없음`);
    this.restoreCalls.push(id);
    this.restoreState(structuredClone(snapshot));
  }
  discardSnapshot(id: number) { this.snapshots.delete(id); }
  override replaceTextInCellDeferredPagination(...args: Parameters<FakeWasm['replaceTextInCellDeferredPagination']>) {
    this.cellTextWrites += 1;
    this.trailingRun = '10pt';
    return super.replaceTextInCellDeferredPagination(...args);
  }
  override borrowDocumentHandle() {
    const base = super.borrowDocumentHandle();
    const tail = () => encoder.encode(`|trailing:${this.trailingRun}`);
    const join = (head: Uint8Array) => {
      const bytes = new Uint8Array(head.length + tail().length);
      bytes.set(head);
      bytes.set(tail(), head.length);
      return bytes;
    };
    return { exportHwp: () => join(base.exportHwp()), exportHwpx: () => join(base.exportHwpx()) };
  }
}

/** InputHandler + CommandHistory를 흉내 낸다: 성공한 snapshot 실행만 undo stack에 before id와 함께 쌓는다. */
class SnapshotHistoryInput extends FakeInput {
  undoStack: Array<{ type: string; beforeId: number | null; cursorBefore: ReturnType<FakeInput['getCursorPosition']> }> = [];
  hookEnabled = true;
  private snapWasm: SnapshotWasm;

  constructor(wasm: SnapshotWasm, eventBus: EventBus) {
    super(wasm, eventBus);
    this.snapWasm = wasm;
  }

  override async executeDocumentAgentOperation(
    desc: Parameters<DocumentAgentInput['executeDocumentAgentOperation']>[0],
    render: () => Promise<void>,
  ) {
    const cursorBefore = this.getCursorPosition();
    const beforeId = this.snapWasm.saveSnapshot();
    const committedBefore = this.transactions;
    try {
      await super.executeDocumentAgentOperation(desc, render);
    } catch (error) {
      this.snapWasm.discardSnapshot(beforeId);
      throw error;
    }
    if (this.transactions === committedBefore) {
      this.snapWasm.discardSnapshot(beforeId);
      return;
    }
    this.undoStack.push({ type: `snapshot:${desc.operationType}`, beforeId, cursorBefore });
  }

  documentAgentUndoSnapshot(operationType: string) {
    if (!this.hookEnabled) return null;
    const top = this.undoStack.at(-1);
    if (!top || top.type !== `snapshot:${operationType}` || top.beforeId === null) return null;
    return { snapshotId: top.beforeId, cursorBefore: { ...top.cursorBefore } };
  }
}

function snapshotHarness(options: { render?: () => Promise<void> } = {}) {
  const wasm = new SnapshotWasm();
  const eventBus = new EventBus();
  const input = new SnapshotHistoryInput(wasm, eventBus);
  const events: unknown[] = [];
  eventBus.on('document-agent-changed', (event) => events.push(event));
  const controller = new DocumentAgentController({
    wasm,
    input,
    eventBus,
    isDirty: () => true,
    render: options.render ?? (async () => {}),
  });
  wasm.cells[1] = [cellParagraph('기존 첫 문단'), cellParagraph('기존 둘째 문단')];
  return { controller, wasm, input, events, eventBus };
}

function exportText(wasm: FakeWasm) {
  return new TextDecoder().decode(wasm.exportHwp());
}

function revertOf(applied: Awaited<ReturnType<DocumentAgentController['applyFieldCommand']>>) {
  return {
    schemaVersion: 1 as const,
    commandId: applied.commandId,
    expectedDocumentEpoch: applied.documentEpoch,
    expectedChangeSeq: applied.afterChangeSeq,
    expectedAfterDocumentSha256: applied.afterDocumentSha256,
    expectedAfterSha256: applied.afterTextSha256,
  };
}

test('같은 세션 field revert는 apply 스냅숏으로 문단 끝 run까지 apply 전 bytes로 복원한다', async () => {
  const { controller, wasm, input, events } = snapshotHarness();
  const beforeBytes = exportText(wasm);
  const applied = await controller.applyFieldCommand(fieldRegionCommand(controller, wasm, '새 첫 문단\n새 둘째 문단'));
  assert.equal(wasm.trailingRun, '10pt');
  const applySnapshotId = input.undoStack.at(-1)!.beforeId;
  const writesBeforeRevert = wasm.cellTextWrites;

  const reverted = await controller.revertFieldCommand(revertOf(applied));
  assert.deepEqual(wasm.restoreCalls, [applySnapshotId]);
  assert.equal(wasm.cellTextWrites, writesBeforeRevert, 'snapshot 경로는 영역 교체를 하지 않는다');
  assert.equal(exportText(wasm), beforeBytes);
  assert.equal(wasm.trailingRun, '11pt');
  assert.equal(reverted.afterDocumentSha256, applied.beforeDocumentSha256);
  assert.equal(reverted.afterTextSha256, applied.beforeTextSha256);
  assert.equal(reverted.beforeChangeSeq + 1, reverted.afterChangeSeq);
  assert.equal(input.transactions, 2);
  assert.equal(events.length, 2);
  assert.equal(input.undoStack.at(-1)!.type, 'snapshot:field-agent:revert');
  // receipt 형식은 v1 그대로다(host 호환).
  assert.deepEqual(Object.keys(reverted).sort(), Object.keys(applied).sort());
  assert.deepEqual(await controller.revertFieldCommand(revertOf(applied)), reverted, 'replay는 같은 receipt');
});

test('hook이 없는 이전 host는 영역 교체로 되돌리고 스냅숏에만 있는 상태는 되살리지 못한다', async () => {
  const { controller, wasm, input } = snapshotHarness();
  input.hookEnabled = false;
  const beforeBytes = exportText(wasm);
  const before = structuredClone(wasm.cells);
  const applied = await controller.applyFieldCommand(fieldRegionCommand(controller, wasm, '새 첫 문단\n새 둘째 문단'));
  const reverted = await controller.revertFieldCommand(revertOf(applied));
  assert.deepEqual(wasm.restoreCalls, []);
  assert.deepEqual(wasm.cells, before);
  assert.equal(wasm.trailingRun, '10pt');
  assert.notEqual(exportText(wasm), beforeBytes);
  assert.equal(reverted.afterTextSha256, applied.beforeTextSha256);
  assert.equal(input.transactions, 2);
});

test('apply 스냅숏이 축출돼 hook이 null이면 영역 교체로 되돌린다', async () => {
  const { controller, wasm, input } = snapshotHarness();
  const before = structuredClone(wasm.cells);
  const applied = await controller.applyFieldCommand(fieldRegionCommand(controller, wasm, '새 첫 문단\n새 둘째 문단'));
  input.undoStack.at(-1)!.beforeId = null; // history budget eviction → discard
  await controller.revertFieldCommand(revertOf(applied));
  assert.deepEqual(wasm.restoreCalls, []);
  assert.deepEqual(wasm.cells, before);
  assert.equal(input.transactions, 2);
});

test('core 스냅숏 복원이 실패하면 rollback 뒤 영역 교체로 한 번만 commit한다', async () => {
  const { controller, wasm, input, events } = snapshotHarness();
  const before = structuredClone(wasm.cells);
  const applied = await controller.applyFieldCommand(fieldRegionCommand(controller, wasm, '새 첫 문단\n새 둘째 문단'));
  wasm.snapshots.delete(input.undoStack.at(-1)!.beforeId!); // core가 무통보로 잃은 스냅숏
  const reverted = await controller.revertFieldCommand(revertOf(applied));
  assert.deepEqual(wasm.cells, before);
  assert.equal(reverted.beforeChangeSeq + 1, reverted.afterChangeSeq);
  assert.equal(input.transactions, 2);
  assert.equal(events.length, 2);
  assert.equal(input.undoStack.filter(entry => entry.type === 'snapshot:field-agent:revert').length, 1);
});

test('복원 bytes가 apply 전 문서 SHA와 다르면 스냅숏 결과를 버리고 영역 교체로 되돌린다', async () => {
  const { controller, wasm, input } = snapshotHarness();
  const before = structuredClone(wasm.cells);
  const applied = await controller.applyFieldCommand(fieldRegionCommand(controller, wasm, '새 첫 문단\n새 둘째 문단'));
  const applySnapshot = wasm.snapshots.get(input.undoStack.at(-1)!.beforeId!)!;
  applySnapshot.trailingRun = '다른 문서';
  const reverted = await controller.revertFieldCommand(revertOf(applied));
  assert.equal(wasm.restoreCalls.length, 1);
  assert.notEqual(wasm.trailingRun, '다른 문서', '불일치 스냅숏 상태는 commit되지 않는다');
  assert.deepEqual(wasm.cells, before);
  assert.equal(reverted.beforeChangeSeq + 1, reverted.afterChangeSeq);
  assert.equal(input.transactions, 2);
});

test('apply 뒤 다른 편집이 있으면 스냅숏 경로도 복원 없이 거부한다', async () => {
  const { controller, wasm, input, eventBus } = snapshotHarness();
  const applied = await controller.applyFieldCommand(fieldRegionCommand(controller, wasm, '새 첫 문단\n새 둘째 문단'));
  wasm.cells[0][0].text = '사용자 편집';
  eventBus.emit('document-mutated', 'user-edit');
  await assert.rejects(
    controller.revertFieldCommand(revertOf(applied)),
    (error: unknown) => (error as { code?: string }).code === 'COMMAND_NOT_LATEST',
  );
  assert.deepEqual(wasm.restoreCalls, []);
  assert.equal(input.transactions, 1);
});

test('다른 snapshot 명령이 undo top이면 apply 스냅숏을 쓰지 않는다', async () => {
  const { controller, wasm, input } = snapshotHarness();
  const applied = await controller.applyFieldCommand(fieldRegionCommand(controller, wasm, '새 첫 문단\n새 둘째 문단'));
  // changeSeq를 바꾸지 않고 top만 바뀐 비정상 상태(방어 확인): hook은 null, 영역 교체로 간다.
  input.undoStack.push({ type: 'snapshot:pasteInternal', beforeId: wasm.saveSnapshot(), cursorBefore: input.getCursorPosition() });
  await controller.revertFieldCommand(revertOf(applied));
  assert.deepEqual(wasm.restoreCalls, []);
});

test('반복 apply→revert도 매번 apply 전 bytes로 돌아간다', async () => {
  const { controller, wasm, input } = snapshotHarness();
  const beforeBytes = exportText(wasm);
  for (let round = 0; round < 3; round += 1) {
    const command = { ...fieldRegionCommand(controller, wasm, `새 첫 문단 ${round}\n새 둘째 문단`), commandId: `region-${round}` };
    const applied = await controller.applyFieldCommand(command);
    await controller.revertFieldCommand(revertOf(applied));
    assert.equal(exportText(wasm), beforeBytes, `round ${round}`);
  }
  assert.equal(wasm.restoreCalls.length, 3);
  assert.equal(input.transactions, 6);
});

test('스냅숏 revert의 strict render 실패는 apply 상태로 복구하고 재시도할 수 있다', async () => {
  let failRender = false;
  const { controller, wasm, input } = snapshotHarness({
    render: async () => { if (failRender) throw new Error('renderer unavailable'); },
  });
  const beforeBytes = exportText(wasm);
  const applied = await controller.applyFieldCommand(fieldRegionCommand(controller, wasm, '새 첫 문단\n새 둘째 문단'));
  const appliedBytes = exportText(wasm);
  failRender = true;
  await assert.rejects(
    controller.revertFieldCommand(revertOf(applied)),
    (error: unknown) => (error as { code?: string; recovered?: boolean }).code === 'RENDER_FAILED',
  );
  assert.equal(exportText(wasm), appliedBytes);
  assert.equal(input.transactions, 1);
  failRender = false;
  await controller.revertFieldCommand(revertOf(applied));
  assert.equal(exportText(wasm), beforeBytes);
});

// ── 5차: 문단 끝 글자 모양(글 뒤 빈 run) 복원 ──

type EndCellParagraph = CellParagraph & { endCharShapeId?: number };

/**
 * 실제 core처럼 문단 끝 위치(offset = 길이)의 글자 모양을 따로 갖는 fake.
 * - 글이 빈 문단에 새 글을 넣으면 끝 모양이 driftEndId로 바뀐다(실제 core의 45 끝 run 재현)
 * - 문단 끝까지 지우면 지운 첫 글자의 모양이 끝 모양으로 남는다
 * - 끝에 넣은 글자는 끝 모양을 이어받는다
 */
class EndShapeWasm extends FakeWasm {
  driftEndId = 45;
  constructor() {
    super();
    this.charShapeProperties.set(21, { textColor: '#000000', italic: false });
    this.charShapeProperties.set(45, { textColor: '#000000', italic: false });
  }
  private paragraphAt(cellIndex: number, cellPara: number) {
    return this.cells[cellIndex][cellPara] as EndCellParagraph;
  }
  endOf(cellIndex: number, cellPara: number) {
    const paragraph = this.paragraphAt(cellIndex, cellPara);
    return paragraph.endCharShapeId ?? paragraph.charShapeIds[Array.from(paragraph.text).length - 1] ?? paragraph.charShapeIds[0];
  }
  override getCellCharPropertiesAt(_section: number, _parentPara: number, _controlIndex: number, cellIndex: number, cellPara: number, offset: number) {
    const paragraph = this.paragraphAt(cellIndex, cellPara);
    const charShapeId = offset >= Array.from(paragraph.text).length
      ? this.endOf(cellIndex, cellPara)
      : paragraph.charShapeIds[offset]!;
    return { charShapeId, ...this.charShapeProperties.get(charShapeId) };
  }
  override replaceTextInCellDeferredPagination(_section: number, _parentPara: number, _controlIndex: number, cellIndex: number, cellPara: number, offset: number, deleteCount: number, text: string) {
    const paragraph = this.paragraphAt(cellIndex, cellPara);
    const chars = Array.from(paragraph.text);
    const length = chars.length;
    const end = this.endOf(cellIndex, cellPara);
    const ids = length === 0 ? [] : paragraph.charShapeIds.slice(0, length);
    let nextEnd: number | undefined = paragraph.endCharShapeId;
    if (deleteCount > 0 && offset + deleteCount === length) nextEnd = ids[offset];
    const inserted = Array.from(text);
    const insertedId = offset < length ? ids[offset]! : end;
    chars.splice(offset, deleteCount, ...inserted);
    ids.splice(offset, deleteCount, ...inserted.map(() => insertedId));
    if (length - deleteCount === 0 && inserted.length > 0) nextEnd = this.driftEndId;
    paragraph.text = chars.join('');
    paragraph.charShapeIds = ids.length > 0 ? ids : [nextEnd ?? end];
    if (nextEnd === undefined) delete paragraph.endCharShapeId;
    else paragraph.endCharShapeId = nextEnd;
    return { ok: true, charOffset: offset + inserted.length, paginationDeferred: true };
  }
}

function endShapeHarness(withSnapshotHook = false) {
  const wasm = new EndShapeWasm();
  const eventBus = new EventBus();
  const input = withSnapshotHook ? null : new FakeInput(wasm, eventBus);
  const controller = new DocumentAgentController({
    wasm,
    input: input!,
    eventBus,
    isDirty: () => true,
    render: async () => {},
  });
  const first = cellParagraph('기존 첫 문단');
  (first as EndCellParagraph).endCharShapeId = 21; // 글 뒤 빈 run 11pt
  const second = cellParagraph('기존 둘째');
  wasm.cells[1] = [first, second];
  return { controller, wasm, input: input!, eventBus };
}

function cellShape(wasm: EndShapeWasm) {
  return wasm.cells[1].map((paragraph, index) => ({
    text: paragraph.text,
    ids: paragraph.charShapeIds.slice(0, Array.from(paragraph.text).length),
    paraShapeId: paragraph.paraShapeId,
    end: wasm.endOf(1, index),
  }));
}

test('영역 evidence는 문단 끝 글자 모양을 담지만 v1 formatSha256은 바꾸지 않는다', () => {
  const { wasm } = endShapeHarness();
  const target = { kind: 'table_cell_region' as const, section: 0, parentPara: 1, controlIndex: 0, cellIndex: 1 };
  const evidence = collectFieldTargetEvidence(wasm, target);
  assert.deepEqual(evidence.regionParagraphFormats?.map(format => format.endCharShapeId), [21, 7]);
  (wasm.cells[1][0] as EndCellParagraph).endCharShapeId = 45;
  assert.equal(collectFieldTargetEvidence(wasm, target).formatSha256, evidence.formatSha256);
});

test('스냅숏 없이 영역 교체로 되돌려도 문단 끝 글자 모양까지 반영 전과 같다', async () => {
  const { controller, wasm, input } = endShapeHarness();
  const before = cellShape(wasm);
  const applied = await controller.applyFieldCommand(fieldRegionCommand(controller, wasm, '새 첫 문단\n새 둘째 문단\n셋째'));
  assert.notEqual(wasm.endOf(1, 0), 21);
  await controller.revertFieldCommand(revertOf(applied));
  assert.deepEqual(cellShape(wasm), before);
  assert.equal(input.transactions, 2);
});

test('끝 모양 정보가 없는 restore-exact(이전 host 계약)는 글 · 글자 모양만 맞고 끝 모양은 부분 복원이다', async () => {
  const { controller, wasm } = endShapeHarness();
  const before = cellShape(wasm);
  const target = { kind: 'table_cell_region' as const, section: 0, parentPara: 1, controlIndex: 0, cellIndex: 1 };
  const beforeEvidence = collectFieldTargetEvidence(wasm, target);
  await controller.applyFieldCommand(fieldRegionCommand(controller, wasm, '새 첫 문단'));
  const state = controller.getDocumentState();
  const current = collectFieldTargetEvidence(wasm, target);
  await controller.applyFieldCommand({
    schemaVersion: 1,
    commandId: 'restore-without-end',
    expectedDocumentEpoch: state.documentEpoch,
    expectedChangeSeq: state.changeSeq,
    expectedDocumentSha256: state.documentSha256,
    target,
    expectedBeforeSha256: current.textSha256,
    expectedFormatSha256: current.formatSha256,
    expectedAdjacentContextSha256: current.adjacentContextSha256,
    replacement: before.map(paragraph => paragraph.text).join('\n'),
    replacementStyle: 'restore-exact',
    replacementFormat: {
      kind: 'table_cell_region',
      paragraphs: before.map(paragraph => ({ length: paragraph.ids.length, charShapeIds: paragraph.ids, paraShapeId: paragraph.paraShapeId })),
    },
    expectedReplacementFormatSha256: beforeEvidence.formatSha256,
  });
  const after = cellShape(wasm);
  assert.deepEqual(after.map(({ end: _end, ...rest }) => rest), before.map(({ end: _end, ...rest }) => rest));
  assert.notEqual(after[0]!.end, 21, '끝 모양 정보가 없으면 되살리지 못한다(새 세션 제한)');
});

test('새 세션 restore-exact에 endCharShapeId가 있으면 문단 끝 글자 모양까지 되살린다', async () => {
  const { controller, wasm } = endShapeHarness();
  const before = cellShape(wasm);
  const target = { kind: 'table_cell_region' as const, section: 0, parentPara: 1, controlIndex: 0, cellIndex: 1 };
  const beforeEvidence = collectFieldTargetEvidence(wasm, target);
  await controller.applyFieldCommand(fieldRegionCommand(controller, wasm, '새 첫 문단'));
  const reloadedBus = new EventBus();
  const reloaded = new DocumentAgentController({
    wasm,
    input: new FakeInput(wasm, reloadedBus),
    eventBus: reloadedBus,
    isDirty: () => true,
    render: async () => {},
  });
  const state = reloaded.getDocumentState();
  const current = collectFieldTargetEvidence(wasm, target);
  await reloaded.applyFieldCommand({
    schemaVersion: 1,
    commandId: 'restore-with-end',
    expectedDocumentEpoch: state.documentEpoch,
    expectedChangeSeq: state.changeSeq,
    expectedDocumentSha256: state.documentSha256,
    target,
    expectedBeforeSha256: current.textSha256,
    expectedFormatSha256: current.formatSha256,
    expectedAdjacentContextSha256: current.adjacentContextSha256,
    replacement: before.map(paragraph => paragraph.text).join('\n'),
    replacementStyle: 'restore-exact',
    replacementFormat: {
      kind: 'table_cell_region',
      paragraphs: before.map(paragraph => ({
        length: paragraph.ids.length,
        charShapeIds: paragraph.ids,
        paraShapeId: paragraph.paraShapeId,
        endCharShapeId: paragraph.end,
      })),
    },
    expectedReplacementFormatSha256: beforeEvidence.formatSha256,
  });
  assert.deepEqual(cellShape(wasm), before);
});
