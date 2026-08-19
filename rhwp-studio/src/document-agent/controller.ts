import { sha256 } from '@noble/hashes/sha2.js';
import { bytesToHex } from '@noble/hashes/utils.js';

import type { EventBus } from '../core/event-bus.ts';
import type { DocumentPosition, FieldInfoResult } from '../core/types.ts';
import {
  DocumentAgentError,
  isDocumentAgentError,
  type RhwpApplyFieldCommandV1,
  type RhwpApplyTextCommandV1,
  type RhwpBodyParagraphTargetV1,
  type RhwpDocumentStateV1,
  type RhwpFieldCommandReceiptV1,
  type RhwpFieldSelectionContextV1,
  type RhwpRevertFieldCommandV1,
  type RhwpRevertTextCommandV1,
  type RhwpSelectionContextV1,
  type RhwpTableCellTextTargetV1,
  type RhwpTextCommandReceiptV1,
} from './types.ts';

const MAX_PARAGRAPH_LENGTH = 4000;
const COMMAND_BUDGET_MS = 3000;
const encoder = new TextEncoder();

type ExportableDocument = {
  exportHwp(): Uint8Array;
  exportHwpx(): Uint8Array;
};

export interface DocumentAgentWasm {
  readonly documentGeneration: number;
  readonly pageCount: number;
  getSourceFormat(): string;
  getSectionCount(): number;
  getParagraphCount(section: number): number;
  getParagraphLength(section: number, paragraph: number): number;
  getTextRange(section: number, paragraph: number, offset: number, count: number): string;
  getControlTextPositions(section: number, paragraph: number): number[];
  getCharPropertiesAt(section: number, paragraph: number, offset: number): { charShapeId?: number };
  getParaPropertiesAt(section: number, paragraph: number): { paraShapeId?: number };
  getStyleAt(section: number, paragraph: number): { id: number; name: string };
  getFieldInfoAt(position: DocumentPosition): FieldInfoResult;
  getTextInCell(
    section: number,
    parentPara: number,
    controlIndex: number,
    cellIndex: number,
    cellParagraph: number,
    offset: number,
    count: number,
  ): string;
  getCellParagraphLength(
    section: number,
    parentPara: number,
    controlIndex: number,
    cellIndex: number,
    cellParagraph: number,
  ): number;
  getCellParagraphCount(
    section: number,
    parentPara: number,
    controlIndex: number,
    cellIndex: number,
  ): number;
  getCellCharPropertiesAt(
    section: number,
    parentPara: number,
    controlIndex: number,
    cellIndex: number,
    cellParagraph: number,
    offset: number,
  ): { charShapeId?: number };
  getCellParaPropertiesAt(
    section: number,
    parentPara: number,
    controlIndex: number,
    cellIndex: number,
    cellParagraph: number,
  ): { paraShapeId?: number };
  getCellOwnProperties(
    section: number,
    parentPara: number,
    controlIndex: number,
    cellIndex: number,
  ): object;
  getTableDimensions(
    section: number,
    parentPara: number,
    controlIndex: number,
  ): { rowCount: number; colCount: number; cellCount: number };
  replaceText(
    section: number,
    paragraph: number,
    offset: number,
    length: number,
    text: string,
  ): { ok: boolean; charOffset?: number; newLength?: number };
  setCharShapeId(
    section: number,
    paragraph: number,
    start: number,
    end: number,
    charShapeId: number,
  ): string;
  setParaShapeId(section: number, paragraph: number, paraShapeId: number): string;
  replaceTextInCellDeferredPagination(
    section: number,
    parentPara: number,
    controlIndex: number,
    cellIndex: number,
    cellParagraph: number,
    offset: number,
    deleteCount: number,
    text: string,
  ): { ok: boolean; charOffset: number; paginationDeferred: boolean };
  insertTextInCellDeferredPagination?(
    section: number,
    parentPara: number,
    controlIndex: number,
    cellIndex: number,
    cellParagraph: number,
    offset: number,
    text: string,
  ): { ok: boolean; charOffset: number; paginationDeferred: boolean };
  deleteTextInCellDeferredPagination?(
    section: number,
    parentPara: number,
    controlIndex: number,
    cellIndex: number,
    cellParagraph: number,
    offset: number,
    count: number,
  ): { ok: boolean; charOffset: number; paginationDeferred: boolean };
  setCharShapeIdInCell(
    section: number,
    parentPara: number,
    controlIndex: number,
    cellIndex: number,
    cellParagraph: number,
    start: number,
    end: number,
    charShapeId: number,
  ): string;
  setCellParaShapeId(
    section: number,
    parentPara: number,
    controlIndex: number,
    cellIndex: number,
    cellParagraph: number,
    paraShapeId: number,
  ): string;
  getPageOfPosition(section: number, paragraph: number): { ok: boolean; page?: number };
  exportHwp(): Uint8Array;
  exportHwpx(): Uint8Array;
  borrowDocumentHandle(): ExportableDocument | null;
  beginDeferredPagination?(): void;
  flushDeferredPagination?(): void;
  cancelDeferredPagination?(): void;
}

export interface DocumentAgentInput {
  getCursorPosition(): DocumentPosition;
  getSelection(): { start: DocumentPosition; end: DocumentPosition } | null;
  executeDocumentAgentOperation(desc: {
    kind: 'snapshot';
    operationType: string;
    operation: (wasm: DocumentAgentWasm) => DocumentPosition | null;
    meta?: {
      actionId?: string;
      domain?: 'text';
      refresh?: 'full';
      dirtyScope?: 'paragraph';
      selection?: 'moveToResult';
    };
  }, render: () => Promise<void>): Promise<void>;
  focusBodyParagraph(section: number, paragraph: number, length: number): boolean;
  focusTableCellText(
    section: number,
    parentPara: number,
    controlIndex: number,
    cellIndex: number,
    cellParagraph: number,
  ): { focused: boolean; page: number };
}

export interface TargetEvidence {
  text: string;
  textSha256: string;
  formatSha256: string;
  adjacentContextSha256: string;
  charShapeId: number;
  paraShapeId: number;
  styleId: number;
}

interface AgentJournalEntry {
  command: RhwpApplyTextCommandV1;
  applyBindingSha256: string;
  applyReceipt: RhwpTextCommandReceiptV1;
  beforeText: string;
  beforeTarget: RhwpBodyParagraphTargetV1;
  afterTarget: RhwpBodyParagraphTargetV1;
  beforeEvidence: TargetEvidence;
  afterEvidence: TargetEvidence;
  nonTargetManifestSha256: string;
  status: 'applied' | 'reverted';
  revertBindingSha256?: string;
  revertReceipt?: RhwpTextCommandReceiptV1;
}

interface FieldAgentJournalEntry {
  command: RhwpApplyFieldCommandV1;
  applyBindingSha256: string;
  applyReceipt: RhwpFieldCommandReceiptV1;
  beforeText: string;
  beforeEvidence: TargetEvidence;
  afterEvidence: TargetEvidence;
  nonTargetManifestSha256: string;
  status: 'applied' | 'reverted';
  revertBindingSha256?: string;
  revertReceipt?: RhwpFieldCommandReceiptV1;
}

interface DocumentAgentControllerDeps {
  wasm: DocumentAgentWasm;
  input: DocumentAgentInput;
  eventBus: EventBus;
  isDirty(): boolean;
  render(): Promise<void>;
  now?: () => number;
}

function digestBytes(value: Uint8Array): string {
  return bytesToHex(sha256(value));
}

function digestText(value: string): string {
  return digestBytes(encoder.encode(value));
}

function codePointLength(value: string): number {
  return Array.from(value).length;
}

/**
 * IME 조합용 atomic replace는 양쪽 텍스트가 1~8자일 때만 허용된다. 필드 제안은
 * 그보다 길 수 있으므로 같은 deferred-pagination snapshot 안에서 기존 전체 값을
 * 삭제한 뒤 새 값을 삽입한다. 두 단계 모두 끝난 뒤 caller가 pagination을 한 번만
 * flush하고 exact postimage를 검증한다.
 */
function replaceWholeFieldTextDeferred(
  wasm: DocumentAgentWasm,
  target: RhwpTableCellTextTargetV1,
  beforeText: string,
  replacement: string,
): { ok: boolean; charOffset: number; paginationDeferred: boolean } {
  const deleteCount = codePointLength(beforeText);
  const replacementLength = codePointLength(replacement);
  if (wasm.deleteTextInCellDeferredPagination && wasm.insertTextInCellDeferredPagination) {
    if (deleteCount > 0) {
      const deleted = wasm.deleteTextInCellDeferredPagination(
        target.section,
        target.parentPara,
        target.controlIndex,
        target.cellIndex,
        target.cellParagraph,
        0,
        deleteCount,
      );
      if (!deleted.ok || deleted.charOffset !== 0) return deleted;
    }
    if (replacementLength > 0) {
      return wasm.insertTextInCellDeferredPagination(
        target.section,
        target.parentPara,
        target.controlIndex,
        target.cellIndex,
        target.cellParagraph,
        0,
        replacement,
      );
    }
    return { ok: true, charOffset: 0, paginationDeferred: true };
  }
  return wasm.replaceTextInCellDeferredPagination(
    target.section,
    target.parentPara,
    target.controlIndex,
    target.cellIndex,
    target.cellParagraph,
    0,
    deleteCount,
    replacement,
  );
}

function safeId(value: number | undefined, label: string): number {
  if (!Number.isSafeInteger(value) || (value as number) < 0) {
    throw new DocumentAgentError('TARGET_FORMAT_MISMATCH', `${label}을 확인할 수 없습니다.`);
  }
  return value as number;
}

function assertTargetCoordinates(wasm: DocumentAgentWasm, target: RhwpBodyParagraphTargetV1): void {
  if (target.kind !== 'body_paragraph'
      || target.charOffset !== 0
      || !Number.isSafeInteger(target.section)
      || !Number.isSafeInteger(target.paragraph)
      || !Number.isSafeInteger(target.length)
      || target.section < 0
      || target.paragraph < 0
      || target.length < 0
      || target.length > MAX_PARAGRAPH_LENGTH
      || target.section >= wasm.getSectionCount()
      || target.paragraph >= wasm.getParagraphCount(target.section)
      || wasm.getParagraphLength(target.section, target.paragraph) !== target.length) {
    throw new DocumentAgentError('TARGET_NOT_FOUND', 'exact body paragraph target을 찾을 수 없습니다.');
  }
}

function charShapeRuns(wasm: DocumentAgentWasm, section: number, paragraph: number, length: number): number[] {
  const ids: number[] = [];
  const count = Math.max(length, 1);
  for (let offset = 0; offset < count; offset += 1) {
    ids.push(safeId(
      wasm.getCharPropertiesAt(section, paragraph, offset).charShapeId,
      'charShapeId',
    ));
  }
  return ids;
}

function paragraphSemantic(
  wasm: DocumentAgentWasm,
  section: number,
  paragraph: number,
): Record<string, unknown> {
  const length = wasm.getParagraphLength(section, paragraph);
  const text = length > 0 ? wasm.getTextRange(section, paragraph, 0, length) : '';
  return {
    section,
    paragraph,
    length,
    textSha256: digestText(text),
    paraShapeId: safeId(
      wasm.getParaPropertiesAt(section, paragraph).paraShapeId,
      'paraShapeId',
    ),
    styleId: safeId(wasm.getStyleAt(section, paragraph).id, 'styleId'),
    charShapeIds: charShapeRuns(wasm, section, paragraph, length),
    controls: wasm.getControlTextPositions(section, paragraph),
  };
}

function adjacentContextSha256(
  wasm: DocumentAgentWasm,
  target: RhwpBodyParagraphTargetV1,
): string {
  const previous = target.paragraph > 0
    ? paragraphSemantic(wasm, target.section, target.paragraph - 1)
    : null;
  const next = target.paragraph + 1 < wasm.getParagraphCount(target.section)
    ? paragraphSemantic(wasm, target.section, target.paragraph + 1)
    : null;
  return digestText(JSON.stringify({ schemaVersion: 1, previous, next }));
}

function hasField(
  wasm: DocumentAgentWasm,
  target: RhwpBodyParagraphTargetV1,
): boolean {
  for (let offset = 0; offset <= target.length; offset += 1) {
    if (wasm.getFieldInfoAt({
      sectionIndex: target.section,
      paragraphIndex: target.paragraph,
      charOffset: offset,
    }).inField) return true;
  }
  return false;
}

export function collectTargetEvidence(
  wasm: DocumentAgentWasm,
  target: RhwpBodyParagraphTargetV1,
): TargetEvidence {
  assertTargetCoordinates(wasm, target);
  if (wasm.getControlTextPositions(target.section, target.paragraph).length > 0 || hasField(wasm, target)) {
    throw new DocumentAgentError(
      'TARGET_FORMAT_MISMATCH',
      '컨트롤 또는 필드가 포함된 문단은 에이전트 명령으로 편집할 수 없습니다.',
    );
  }
  const charShapeIds = charShapeRuns(wasm, target.section, target.paragraph, target.length);
  const charShapeId = charShapeIds[0];
  if (!charShapeIds.every(id => id === charShapeId)) {
    throw new DocumentAgentError(
      'TARGET_FORMAT_MISMATCH',
      '혼합 글자 서식 문단은 에이전트 명령으로 편집할 수 없습니다.',
    );
  }
  const paraShapeId = safeId(
    wasm.getParaPropertiesAt(target.section, target.paragraph).paraShapeId,
    'paraShapeId',
  );
  const styleId = safeId(wasm.getStyleAt(target.section, target.paragraph).id, 'styleId');
  const text = target.length > 0
    ? wasm.getTextRange(target.section, target.paragraph, 0, target.length)
    : '';
  return {
    text,
    textSha256: digestText(text),
    formatSha256: digestText(JSON.stringify({
      schemaVersion: 1,
      charShapeId,
      paraShapeId,
      styleId,
    })),
    adjacentContextSha256: adjacentContextSha256(wasm, target),
    charShapeId,
    paraShapeId,
    styleId,
  };
}

function stableJson(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(stableJson).join(',')}]`;
  if (value && typeof value === 'object') {
    return `{${Object.entries(value as Record<string, unknown>)
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([key, entry]) => `${JSON.stringify(key)}:${stableJson(entry)}`)
      .join(',')}}`;
  }
  return JSON.stringify(value) ?? 'null';
}

function assertFieldTargetCoordinates(
  wasm: DocumentAgentWasm,
  target: RhwpTableCellTextTargetV1,
): number {
  try {
    const dimensions = wasm.getTableDimensions(
      target.section,
      target.parentPara,
      target.controlIndex,
    );
    if (!Number.isSafeInteger(dimensions.cellCount)
        || target.cellIndex >= dimensions.cellCount) throw new Error('cell index');
    const paragraphCount = wasm.getCellParagraphCount(
      target.section,
      target.parentPara,
      target.controlIndex,
      target.cellIndex,
    );
    if (!Number.isSafeInteger(paragraphCount)
        || target.cellParagraph >= paragraphCount) throw new Error('cell paragraph');
    const length = wasm.getCellParagraphLength(
      target.section,
      target.parentPara,
      target.controlIndex,
      target.cellIndex,
      target.cellParagraph,
    );
    if (!Number.isSafeInteger(length) || length < 0 || length > MAX_PARAGRAPH_LENGTH) {
      throw new Error('cell paragraph length');
    }
    return length;
  } catch {
    throw new DocumentAgentError('TARGET_NOT_FOUND', 'exact table cell text target을 찾을 수 없습니다.');
  }
}

function cellCharShapeRuns(
  wasm: DocumentAgentWasm,
  target: RhwpTableCellTextTargetV1,
  length: number,
): number[] {
  const ids: number[] = [];
  for (let offset = 0; offset < Math.max(length, 1); offset += 1) {
    ids.push(safeId(wasm.getCellCharPropertiesAt(
      target.section,
      target.parentPara,
      target.controlIndex,
      target.cellIndex,
      target.cellParagraph,
      offset,
    ).charShapeId, 'cell charShapeId'));
  }
  return ids;
}

function cellParagraphSemantic(
  wasm: DocumentAgentWasm,
  target: RhwpTableCellTextTargetV1,
  cellIndex: number,
  cellParagraph: number,
): Record<string, unknown> {
  const length = wasm.getCellParagraphLength(
    target.section,
    target.parentPara,
    target.controlIndex,
    cellIndex,
    cellParagraph,
  );
  const paragraphTarget = { ...target, cellIndex, cellParagraph };
  const text = length > 0 ? wasm.getTextInCell(
    target.section,
    target.parentPara,
    target.controlIndex,
    cellIndex,
    cellParagraph,
    0,
    length,
  ) : '';
  return {
    cellIndex,
    cellParagraph,
    length,
    textSha256: digestText(text),
    paraShapeId: safeId(wasm.getCellParaPropertiesAt(
      target.section,
      target.parentPara,
      target.controlIndex,
      cellIndex,
      cellParagraph,
    ).paraShapeId, 'cell paraShapeId'),
    charShapeIds: cellCharShapeRuns(wasm, paragraphTarget, length),
  };
}

function fieldNonTargetManifestSha256(
  wasm: DocumentAgentWasm,
  target: RhwpTableCellTextTargetV1,
): string {
  const dimensions = wasm.getTableDimensions(
    target.section,
    target.parentPara,
    target.controlIndex,
  );
  const cells: Array<Record<string, unknown>> = [];
  for (let cellIndex = 0; cellIndex < dimensions.cellCount; cellIndex += 1) {
    const paragraphCount = wasm.getCellParagraphCount(
      target.section,
      target.parentPara,
      target.controlIndex,
      cellIndex,
    );
    const paragraphs: Array<Record<string, unknown>> = [];
    for (let cellParagraph = 0; cellParagraph < paragraphCount; cellParagraph += 1) {
      if (cellIndex === target.cellIndex && cellParagraph === target.cellParagraph) continue;
      paragraphs.push(cellParagraphSemantic(wasm, target, cellIndex, cellParagraph));
    }
    cells.push({
      cellIndex,
      properties: wasm.getCellOwnProperties(
        target.section,
        target.parentPara,
        target.controlIndex,
        cellIndex,
      ),
      paragraphCount,
      paragraphs,
    });
  }
  return digestText(stableJson({
    schemaVersion: 1,
    table: {
      section: target.section,
      parentPara: target.parentPara,
      controlIndex: target.controlIndex,
      dimensions,
      cells,
    },
  }));
}

/** 서버와 SDK가 같은 exact cell preimage를 결속할 수 있도록 공개하는 field evidence. */
export function collectFieldTargetEvidence(
  wasm: DocumentAgentWasm,
  target: RhwpTableCellTextTargetV1,
): TargetEvidence {
  const length = assertFieldTargetCoordinates(wasm, target);
  const charShapeIds = cellCharShapeRuns(wasm, target, length);
  const charShapeId = charShapeIds[0];
  if (!charShapeIds.every(id => id === charShapeId)) {
    throw new DocumentAgentError(
      'TARGET_FORMAT_MISMATCH',
      '혼합 글자 서식 셀 문단은 에이전트 명령으로 편집할 수 없습니다.',
    );
  }
  const paraShapeId = safeId(wasm.getCellParaPropertiesAt(
    target.section,
    target.parentPara,
    target.controlIndex,
    target.cellIndex,
    target.cellParagraph,
  ).paraShapeId, 'cell paraShapeId');
  const text = length > 0 ? wasm.getTextInCell(
    target.section,
    target.parentPara,
    target.controlIndex,
    target.cellIndex,
    target.cellParagraph,
    0,
    length,
  ) : '';
  const adjacentContextSha256 = fieldNonTargetManifestSha256(wasm, target);
  return {
    text,
    textSha256: digestText(text),
    formatSha256: digestText(stableJson({
      schemaVersion: 1,
      charShapeId,
      paraShapeId,
      cellProperties: wasm.getCellOwnProperties(
        target.section,
        target.parentPara,
        target.controlIndex,
        target.cellIndex,
      ),
    })),
    adjacentContextSha256,
    charShapeId: charShapeId!,
    paraShapeId,
    styleId: 0,
  };
}

function nonTargetManifestSha256(
  wasm: DocumentAgentWasm,
  target: RhwpBodyParagraphTargetV1,
): string {
  const paragraphs: Array<Record<string, unknown>> = [];
  for (let section = 0; section < wasm.getSectionCount(); section += 1) {
    for (let paragraph = 0; paragraph < wasm.getParagraphCount(section); paragraph += 1) {
      if (section === target.section && paragraph === target.paragraph) continue;
      paragraphs.push(paragraphSemantic(wasm, section, paragraph));
    }
  }
  return digestText(JSON.stringify({
    schemaVersion: 1,
    sectionCount: wasm.getSectionCount(),
    paragraphCounts: Array.from(
      { length: wasm.getSectionCount() },
      (_, section) => wasm.getParagraphCount(section),
    ),
    paragraphs,
  }));
}

function exportDocumentSha256(wasm: DocumentAgentWasm, format: 'hwp' | 'hwpx'): string {
  const document = wasm.borrowDocumentHandle();
  if (!document) throw new DocumentAgentError('TARGET_NOT_FOUND', '문서가 로드되지 않았습니다.');
  return digestBytes(format === 'hwpx' ? document.exportHwpx() : document.exportHwp());
}

function exportPersistedDocumentSha256(
  wasm: DocumentAgentWasm,
  format: 'hwp' | 'hwpx',
): string {
  // SDK export와 같은 WasmBridge 경로를 거쳐 저장 시점 캐럿까지 스탬핑한다.
  // selection-only 이동은 apply preimage fence를 바꾸면 안 되므로 일반 state 읽기는
  // 위 raw export를 쓰고, terminal receipt만 이 persisted export에 결속한다.
  return digestBytes(format === 'hwpx' ? wasm.exportHwpx() : wasm.exportHwp());
}

function commandBinding(command: RhwpApplyTextCommandV1): string {
  return digestText(JSON.stringify({
    schemaVersion: command.schemaVersion,
    commandId: command.commandId,
    expectedDocumentEpoch: command.expectedDocumentEpoch,
    expectedChangeSeq: command.expectedChangeSeq,
    expectedDocumentSha256: command.expectedDocumentSha256,
    target: command.target,
    expectedBeforeSha256: command.expectedBeforeSha256,
    expectedFormatSha256: command.expectedFormatSha256,
    expectedAdjacentContextSha256: command.expectedAdjacentContextSha256,
    replacement: command.replacement,
  }));
}

function fieldCommandBinding(command: RhwpApplyFieldCommandV1): string {
  return digestText(stableJson(command));
}

function revertBinding(command: RhwpRevertTextCommandV1): string {
  return digestText(JSON.stringify({
    schemaVersion: command.schemaVersion,
    commandId: command.commandId,
    expectedDocumentEpoch: command.expectedDocumentEpoch,
    expectedChangeSeq: command.expectedChangeSeq,
    expectedAfterDocumentSha256: command.expectedAfterDocumentSha256,
    expectedAfterSha256: command.expectedAfterSha256,
  }));
}

function sameBodyPosition(a: DocumentPosition, b: DocumentPosition): boolean {
  return a.parentParaIndex === undefined
    && b.parentParaIndex === undefined
    && a.sectionIndex === b.sectionIndex
    && a.paragraphIndex === b.paragraphIndex;
}

function samePosition(a: DocumentPosition, b: DocumentPosition): boolean {
  return sameBodyPosition(a, b) && a.charOffset === b.charOffset;
}

export class DocumentAgentController {
  private readonly deps: DocumentAgentControllerDeps;
  private readonly now: () => number;
  private documentEpoch: number;
  private changeSeq = 0;
  private latest: AgentJournalEntry | null = null;
  private readonly journal = new Map<string, AgentJournalEntry>();
  private latestField: FieldAgentJournalEntry | null = null;
  private readonly fieldJournal = new Map<string, FieldAgentJournalEntry>();
  private readonly offMutation: () => void;

  constructor(deps: DocumentAgentControllerDeps) {
    this.deps = deps;
    this.now = deps.now ?? (() => performance.now());
    this.documentEpoch = deps.wasm.documentGeneration;
    this.offMutation = deps.eventBus.on('document-mutated', () => {
      this.syncGeneration();
      this.changeSeq += 1;
    });
  }

  dispose(): void {
    this.offMutation();
  }

  getDocumentState(): RhwpDocumentStateV1 {
    this.syncGeneration();
    const format = this.currentFormat();
    return {
      schemaVersion: 1,
      format,
      documentEpoch: this.documentEpoch,
      changeSeq: this.changeSeq,
      dirty: this.deps.isDirty(),
      pageCount: this.deps.wasm.pageCount,
      documentSha256: exportDocumentSha256(this.deps.wasm, format),
    };
  }

  getSelectionContext(): RhwpSelectionContextV1 {
    this.syncGeneration();
    const position = this.deps.input.getCursorPosition();
    const page = this.pageFor(position.sectionIndex, position.paragraphIndex);
    const selection = this.deps.input.getSelection();
    const collapsed = !selection || samePosition(selection.start, selection.end);
    let target: RhwpBodyParagraphTargetV1 | null = null;
    let editable = false;
    let selectedTextSha256: string | null = null;

    if (position.parentParaIndex === undefined
        && position.sectionIndex >= 0
        && position.paragraphIndex >= 0
        && position.sectionIndex < this.deps.wasm.getSectionCount()
        && position.paragraphIndex < this.deps.wasm.getParagraphCount(position.sectionIndex)) {
      const length = this.deps.wasm.getParagraphLength(
        position.sectionIndex,
        position.paragraphIndex,
      );
      if (length <= MAX_PARAGRAPH_LENGTH) {
        target = {
          kind: 'body_paragraph',
          section: position.sectionIndex,
          paragraph: position.paragraphIndex,
          charOffset: 0,
          length,
        };
        try {
          collectTargetEvidence(this.deps.wasm, target);
          this.currentFormat();
          editable = true;
        } catch {
          editable = false;
        }
      }
    }

    if (!collapsed && selection && sameBodyPosition(selection.start, selection.end)) {
      const length = selection.end.charOffset - selection.start.charOffset;
      if (length >= 0) {
        selectedTextSha256 = digestText(this.deps.wasm.getTextRange(
          selection.start.sectionIndex,
          selection.start.paragraphIndex,
          selection.start.charOffset,
          length,
        ));
      }
    }

    return {
      schemaVersion: 1,
      documentEpoch: this.documentEpoch,
      changeSeq: this.changeSeq,
      page,
      editable,
      collapsed,
      target,
      selectedTextSha256,
    };
  }

  getFieldSelectionContext(): RhwpFieldSelectionContextV1 {
    this.syncGeneration();
    const position = this.deps.input.getCursorPosition();
    const parentPara = position.parentParaIndex;
    const controlIndex = position.controlIndex;
    const cellIndex = position.cellIndex;
    const cellParagraph = position.cellParaIndex;
    const page = this.pageFor(
      position.sectionIndex,
      parentPara ?? position.paragraphIndex,
    );
    let target: RhwpTableCellTextTargetV1 | null = null;
    let editable = false;

    if ([position.sectionIndex, parentPara, controlIndex, cellIndex, cellParagraph]
      .every((value) => Number.isSafeInteger(value) && (value as number) >= 0)) {
      target = {
        kind: 'table_cell_text',
        section: position.sectionIndex,
        parentPara: parentPara!,
        controlIndex: controlIndex!,
        cellIndex: cellIndex!,
        cellParagraph: cellParagraph!,
      };
      try {
        collectFieldTargetEvidence(this.deps.wasm, target);
        this.currentFormat();
        editable = true;
      } catch {
        editable = false;
      }
    }

    return {
      schemaVersion: 1,
      documentEpoch: this.documentEpoch,
      changeSeq: this.changeSeq,
      page,
      editable,
      target,
    };
  }

  async applyTextCommand(command: RhwpApplyTextCommandV1): Promise<RhwpTextCommandReceiptV1> {
    this.syncGeneration();
    const binding = commandBinding(command);
    const replay = this.journal.get(command.commandId);
    if (replay) {
      if (replay.status === 'applied'
          && replay.applyBindingSha256 === binding
          && this.isCurrentReceiptState(
            replay.applyReceipt,
            replay.afterTarget,
            replay.applyReceipt.afterTextSha256,
          )) {
        return replay.applyReceipt;
      }
      throw new DocumentAgentError(
        'COMMAND_REPLAY_MISMATCH',
        '같은 commandId가 다른 binding 또는 terminal 상태로 이미 사용되었습니다.',
      );
    }

    const startedAt = this.now();
    const state = this.getDocumentState();
    this.assertApplyFence(command, state);
    const beforeEvidence = collectTargetEvidence(this.deps.wasm, command.target);
    if (beforeEvidence.textSha256 !== command.expectedBeforeSha256) {
      throw new DocumentAgentError('TARGET_PREIMAGE_MISMATCH', 'target before SHA가 다릅니다.');
    }
    if (beforeEvidence.formatSha256 !== command.expectedFormatSha256) {
      throw new DocumentAgentError('TARGET_FORMAT_MISMATCH', 'target format SHA가 다릅니다.');
    }
    if (beforeEvidence.adjacentContextSha256 !== command.expectedAdjacentContextSha256) {
      throw new DocumentAgentError('TARGET_CONTEXT_MISMATCH', 'target adjacent context SHA가 다릅니다.');
    }
    const beforeNonTarget = nonTargetManifestSha256(this.deps.wasm, command.target);
    this.assertWithinBudget(startedAt);

    const afterTarget: RhwpBodyParagraphTargetV1 = {
      ...command.target,
      length: codePointLength(command.replacement),
    };
    let afterEvidence!: TargetEvidence;
    let afterDocumentSha256 = '';
    const beforeSeq = this.changeSeq;

    try {
      await this.deps.input.executeDocumentAgentOperation({
        kind: 'snapshot',
        operationType: 'document-agent:apply',
        operation: (wasm) => {
          this.assertRuntimeFence(command.expectedDocumentEpoch, beforeSeq);
          let deferred = false;
          try {
            wasm.beginDeferredPagination?.();
            deferred = true;
            const result = wasm.replaceText(
              command.target.section,
              command.target.paragraph,
              0,
              command.target.length,
              command.replacement,
            );
            if (!result.ok || result.newLength !== afterTarget.length) {
              throw new DocumentAgentError('TRANSACTION_FAILED', 'replaceText가 target을 교체하지 못했습니다.');
            }
            if (afterTarget.length > 0) {
              wasm.setCharShapeId(
                command.target.section,
                command.target.paragraph,
                0,
                afterTarget.length,
                beforeEvidence.charShapeId,
              );
            }
            wasm.setParaShapeId(
              command.target.section,
              command.target.paragraph,
              beforeEvidence.paraShapeId,
            );
            wasm.flushDeferredPagination?.();
            deferred = false;

            afterEvidence = collectTargetEvidence(wasm, afterTarget);
            if (afterEvidence.text !== command.replacement) {
              throw new DocumentAgentError('TARGET_PREIMAGE_MISMATCH', 'target postimage가 replacement와 다릅니다.');
            }
            if (afterEvidence.formatSha256 !== beforeEvidence.formatSha256) {
              throw new DocumentAgentError('TARGET_FORMAT_MISMATCH', 'target format이 변경되었습니다.');
            }
            if (afterEvidence.adjacentContextSha256 !== beforeEvidence.adjacentContextSha256) {
              throw new DocumentAgentError('TARGET_CONTEXT_MISMATCH', 'adjacent context가 변경되었습니다.');
            }
            if (nonTargetManifestSha256(wasm, afterTarget) !== beforeNonTarget) {
              throw new DocumentAgentError('NON_TARGET_CHANGED', 'target 밖 semantic manifest가 변경되었습니다.');
            }
            if (wasm.pageCount !== state.pageCount) {
              throw new DocumentAgentError('PAGE_COUNT_CHANGED', '문서 페이지 수가 변경되었습니다.');
            }
            afterDocumentSha256 = exportDocumentSha256(wasm, state.format);
            this.assertWithinBudget(startedAt);
            return {
              sectionIndex: afterTarget.section,
              paragraphIndex: afterTarget.paragraph,
              charOffset: afterTarget.length,
            };
          } catch (error) {
            if (deferred) {
              try { wasm.cancelDeferredPagination?.(); } catch { /* snapshot rollback이 최종 복구한다. */ }
            }
            throw error;
          }
        },
        meta: {
          actionId: 'document-agent:apply',
          domain: 'text',
          refresh: 'full',
          dirtyScope: 'paragraph',
          selection: 'moveToResult',
        },
      }, async () => {
        await this.deps.render();
        this.assertWithinBudget(startedAt);
      });
    } catch (error) {
      throw this.normalizeExecutionError(error);
    }

    if (this.changeSeq !== beforeSeq + 1) {
      throw new DocumentAgentError(
        'TRANSACTION_FAILED',
        'changeSeq가 정확히 1 증가하지 않았습니다.',
        false,
      );
    }
    afterDocumentSha256 = exportPersistedDocumentSha256(this.deps.wasm, state.format);
    const receipt: RhwpTextCommandReceiptV1 = {
      schemaVersion: 1,
      commandId: command.commandId,
      operation: 'apply',
      documentEpoch: this.documentEpoch,
      beforeChangeSeq: beforeSeq,
      afterChangeSeq: this.changeSeq,
      beforeDocumentSha256: state.documentSha256,
      afterDocumentSha256,
      beforeTextSha256: beforeEvidence.textSha256,
      afterTextSha256: afterEvidence.textSha256,
      formatSha256: afterEvidence.formatSha256,
      adjacentContextSha256: afterEvidence.adjacentContextSha256,
      pageCountBefore: state.pageCount,
      pageCountAfter: this.deps.wasm.pageCount,
      target: command.target,
    };
    const entry: AgentJournalEntry = {
      command: structuredClone(command),
      applyBindingSha256: binding,
      applyReceipt: receipt,
      beforeText: beforeEvidence.text,
      beforeTarget: command.target,
      afterTarget,
      beforeEvidence,
      afterEvidence,
      nonTargetManifestSha256: beforeNonTarget,
      status: 'applied',
    };
    this.journal.set(command.commandId, entry);
    this.latest = entry;
    this.emitDocumentAgentChanged({
      schemaVersion: 1,
      reason: 'agent_apply',
      documentEpoch: this.documentEpoch,
      changeSeq: this.changeSeq,
      commandId: command.commandId,
    });
    return receipt;
  }

  async revertTextCommand(command: RhwpRevertTextCommandV1): Promise<RhwpTextCommandReceiptV1> {
    this.syncGeneration();
    const binding = revertBinding(command);
    const entry = this.journal.get(command.commandId);
    if (entry?.status === 'reverted') {
      if (entry.revertBindingSha256 === binding
          && entry.revertReceipt
          && this.isCurrentReceiptState(
            entry.revertReceipt,
            entry.beforeTarget,
            entry.revertReceipt.afterTextSha256,
          )) return entry.revertReceipt;
      throw new DocumentAgentError(
        'COMMAND_REPLAY_MISMATCH',
        'revert command binding 또는 현재 terminal receipt 상태가 다릅니다.',
      );
    }
    if (!entry || this.latest !== entry || entry.status !== 'applied'
        || this.changeSeq !== entry.applyReceipt.afterChangeSeq) {
      throw new DocumentAgentError('COMMAND_NOT_LATEST', '가장 최근 exact command만 되돌릴 수 있습니다.');
    }

    const startedAt = this.now();
    const state = this.getDocumentState();
    if (command.expectedDocumentEpoch !== state.documentEpoch) {
      throw new DocumentAgentError('DOCUMENT_EPOCH_MISMATCH', 'document epoch가 다릅니다.');
    }
    if (command.expectedChangeSeq !== state.changeSeq) {
      throw new DocumentAgentError('CHANGE_SEQ_MISMATCH', 'changeSeq가 다릅니다.');
    }
    if (command.expectedAfterDocumentSha256 !== state.documentSha256) {
      throw new DocumentAgentError('DOCUMENT_SHA_MISMATCH', 'after document SHA가 다릅니다.');
    }
    const currentEvidence = collectTargetEvidence(this.deps.wasm, entry.afterTarget);
    if (currentEvidence.textSha256 !== command.expectedAfterSha256
        || currentEvidence.textSha256 !== entry.applyReceipt.afterTextSha256) {
      throw new DocumentAgentError('TARGET_PREIMAGE_MISMATCH', 'revert target after SHA가 다릅니다.');
    }
    if (nonTargetManifestSha256(this.deps.wasm, entry.afterTarget)
        !== entry.nonTargetManifestSha256) {
      throw new DocumentAgentError('COMMAND_NOT_LATEST', 'target 밖 변경이 있어 되돌릴 수 없습니다.');
    }
    this.assertWithinBudget(startedAt);

    let afterDocumentSha256 = '';
    let revertedEvidence!: TargetEvidence;
    const beforeSeq = this.changeSeq;
    try {
      await this.deps.input.executeDocumentAgentOperation({
        kind: 'snapshot',
        operationType: 'document-agent:revert',
        operation: (wasm) => {
          this.assertRuntimeFence(command.expectedDocumentEpoch, beforeSeq);
          let deferred = false;
          try {
            wasm.beginDeferredPagination?.();
            deferred = true;
            const result = wasm.replaceText(
              entry.afterTarget.section,
              entry.afterTarget.paragraph,
              0,
              entry.afterTarget.length,
              entry.beforeText,
            );
            if (!result.ok || result.newLength !== entry.beforeTarget.length) {
              throw new DocumentAgentError('TRANSACTION_FAILED', 'inverse replaceText가 실패했습니다.');
            }
            if (entry.beforeTarget.length > 0) {
              wasm.setCharShapeId(
                entry.beforeTarget.section,
                entry.beforeTarget.paragraph,
                0,
                entry.beforeTarget.length,
                entry.beforeEvidence.charShapeId,
              );
            }
            wasm.setParaShapeId(
              entry.beforeTarget.section,
              entry.beforeTarget.paragraph,
              entry.beforeEvidence.paraShapeId,
            );
            wasm.flushDeferredPagination?.();
            deferred = false;

            revertedEvidence = collectTargetEvidence(wasm, entry.beforeTarget);
            if (revertedEvidence.textSha256 !== entry.beforeEvidence.textSha256) {
              throw new DocumentAgentError('TARGET_PREIMAGE_MISMATCH', 'before target 복원이 일치하지 않습니다.');
            }
            if (revertedEvidence.formatSha256 !== entry.beforeEvidence.formatSha256) {
              throw new DocumentAgentError('TARGET_FORMAT_MISMATCH', 'before format 복원이 일치하지 않습니다.');
            }
            if (revertedEvidence.adjacentContextSha256
                !== entry.beforeEvidence.adjacentContextSha256) {
              throw new DocumentAgentError('TARGET_CONTEXT_MISMATCH', 'before context 복원이 일치하지 않습니다.');
            }
            if (nonTargetManifestSha256(wasm, entry.beforeTarget)
                !== entry.nonTargetManifestSha256) {
              throw new DocumentAgentError('NON_TARGET_CHANGED', 'revert가 target 밖을 변경했습니다.');
            }
            if (wasm.pageCount !== entry.applyReceipt.pageCountBefore) {
              throw new DocumentAgentError('PAGE_COUNT_CHANGED', 'revert 뒤 페이지 수가 다릅니다.');
            }
            afterDocumentSha256 = exportDocumentSha256(wasm, state.format);
            this.assertWithinBudget(startedAt);
            return {
              sectionIndex: entry.beforeTarget.section,
              paragraphIndex: entry.beforeTarget.paragraph,
              charOffset: entry.beforeTarget.length,
            };
          } catch (error) {
            if (deferred) {
              try { wasm.cancelDeferredPagination?.(); } catch { /* snapshot rollback이 최종 복구한다. */ }
            }
            throw error;
          }
        },
        meta: {
          actionId: 'document-agent:revert',
          domain: 'text',
          refresh: 'full',
          dirtyScope: 'paragraph',
          selection: 'moveToResult',
        },
      }, async () => {
        await this.deps.render();
        this.assertWithinBudget(startedAt);
      });
    } catch (error) {
      throw this.normalizeExecutionError(error);
    }

    if (this.changeSeq !== beforeSeq + 1) {
      throw new DocumentAgentError(
        'TRANSACTION_FAILED',
        'revert changeSeq가 정확히 1 증가하지 않았습니다.',
        false,
      );
    }
    afterDocumentSha256 = exportPersistedDocumentSha256(this.deps.wasm, state.format);
    const receipt: RhwpTextCommandReceiptV1 = {
      schemaVersion: 1,
      commandId: command.commandId,
      operation: 'revert',
      documentEpoch: this.documentEpoch,
      beforeChangeSeq: beforeSeq,
      afterChangeSeq: this.changeSeq,
      beforeDocumentSha256: state.documentSha256,
      afterDocumentSha256,
      beforeTextSha256: currentEvidence.textSha256,
      afterTextSha256: revertedEvidence.textSha256,
      formatSha256: revertedEvidence.formatSha256,
      adjacentContextSha256: revertedEvidence.adjacentContextSha256,
      pageCountBefore: state.pageCount,
      pageCountAfter: this.deps.wasm.pageCount,
      target: entry.beforeTarget,
    };
    entry.status = 'reverted';
    entry.revertBindingSha256 = binding;
    entry.revertReceipt = receipt;
    this.emitDocumentAgentChanged({
      schemaVersion: 1,
      reason: 'agent_revert',
      documentEpoch: this.documentEpoch,
      changeSeq: this.changeSeq,
      commandId: command.commandId,
    });
    return receipt;
  }

  async applyFieldCommand(command: RhwpApplyFieldCommandV1): Promise<RhwpFieldCommandReceiptV1> {
    this.syncGeneration();
    const binding = fieldCommandBinding(command);
    const replay = this.fieldJournal.get(command.commandId);
    if (replay) {
      if (replay.status === 'applied'
          && replay.applyBindingSha256 === binding
          && this.isCurrentFieldReceiptState(
            replay.applyReceipt,
            replay.applyReceipt.afterTextSha256,
          )) return replay.applyReceipt;
      throw new DocumentAgentError(
        'COMMAND_REPLAY_MISMATCH',
        '같은 field commandId가 다른 binding 또는 terminal 상태로 이미 사용되었습니다.',
      );
    }

    const startedAt = this.now();
    const state = this.getDocumentState();
    this.assertApplyFence(command, state);
    const beforeEvidence = collectFieldTargetEvidence(this.deps.wasm, command.target);
    if (beforeEvidence.textSha256 !== command.expectedBeforeSha256) {
      throw new DocumentAgentError('TARGET_PREIMAGE_MISMATCH', 'field target before SHA가 다릅니다.');
    }
    if (beforeEvidence.formatSha256 !== command.expectedFormatSha256) {
      throw new DocumentAgentError('TARGET_FORMAT_MISMATCH', 'field target format SHA가 다릅니다.');
    }
    if (beforeEvidence.adjacentContextSha256 !== command.expectedAdjacentContextSha256) {
      throw new DocumentAgentError('TARGET_CONTEXT_MISMATCH', 'field target context SHA가 다릅니다.');
    }
    const beforeNonTarget = beforeEvidence.adjacentContextSha256;
    this.assertWithinBudget(startedAt);

    let afterEvidence!: TargetEvidence;
    let afterDocumentSha256 = '';
    const beforeSeq = this.changeSeq;
    try {
      await this.deps.input.executeDocumentAgentOperation({
        kind: 'snapshot',
        operationType: 'field-agent:apply',
        operation: (wasm) => {
          this.assertRuntimeFence(command.expectedDocumentEpoch, beforeSeq);
          let deferred = false;
          try {
            wasm.beginDeferredPagination?.();
            deferred = true;
            const result = replaceWholeFieldTextDeferred(
              wasm,
              command.target,
              beforeEvidence.text,
              command.replacement,
            );
            if (!result.ok || result.charOffset !== codePointLength(command.replacement)) {
              throw new DocumentAgentError('TRANSACTION_FAILED', 'field target 셀 텍스트를 교체하지 못했습니다.');
            }
            const afterLength = codePointLength(command.replacement);
            if (afterLength > 0) {
              wasm.setCharShapeIdInCell(
                command.target.section,
                command.target.parentPara,
                command.target.controlIndex,
                command.target.cellIndex,
                command.target.cellParagraph,
                0,
                afterLength,
                beforeEvidence.charShapeId,
              );
            }
            wasm.setCellParaShapeId(
              command.target.section,
              command.target.parentPara,
              command.target.controlIndex,
              command.target.cellIndex,
              command.target.cellParagraph,
              beforeEvidence.paraShapeId,
            );
            wasm.flushDeferredPagination?.();
            deferred = false;

            afterEvidence = collectFieldTargetEvidence(wasm, command.target);
            if (afterEvidence.text !== command.replacement) {
              throw new DocumentAgentError('TARGET_PREIMAGE_MISMATCH', 'field target postimage가 replacement와 다릅니다.');
            }
            if (afterEvidence.formatSha256 !== beforeEvidence.formatSha256) {
              throw new DocumentAgentError('TARGET_FORMAT_MISMATCH', 'field target format이 변경되었습니다.');
            }
            if (afterEvidence.adjacentContextSha256 !== beforeNonTarget) {
              throw new DocumentAgentError('NON_TARGET_CHANGED', 'field target 밖 표 내용이 변경되었습니다.');
            }
            if (wasm.pageCount !== state.pageCount) {
              throw new DocumentAgentError('PAGE_COUNT_CHANGED', 'field apply 뒤 페이지 수가 변경되었습니다.');
            }
            afterDocumentSha256 = exportDocumentSha256(wasm, state.format);
            this.assertWithinBudget(startedAt);
            return {
              sectionIndex: command.target.section,
              paragraphIndex: command.target.cellParagraph,
              charOffset: afterLength,
              parentParaIndex: command.target.parentPara,
              controlIndex: command.target.controlIndex,
              cellIndex: command.target.cellIndex,
              cellParaIndex: command.target.cellParagraph,
            };
          } catch (error) {
            if (deferred) {
              try { wasm.cancelDeferredPagination?.(); } catch { /* snapshot rollback이 최종 복구한다. */ }
            }
            throw error;
          }
        },
        meta: {
          actionId: 'field-agent:apply',
          domain: 'text',
          refresh: 'full',
          dirtyScope: 'paragraph',
          selection: 'moveToResult',
        },
      }, async () => {
        await this.deps.render();
        this.assertWithinBudget(startedAt);
      });
    } catch (error) {
      throw this.normalizeExecutionError(error);
    }
    if (this.changeSeq !== beforeSeq + 1) {
      throw new DocumentAgentError('TRANSACTION_FAILED', 'field apply changeSeq가 정확히 1 증가하지 않았습니다.', false);
    }
    // operation 내부 export는 rollback gate이고, public receipt는 SDK가 실제 저장할
    // onBeforeExport 적용 bytes에 결속한다.
    afterDocumentSha256 = exportPersistedDocumentSha256(this.deps.wasm, state.format);

    const receipt: RhwpFieldCommandReceiptV1 = {
      schemaVersion: 1,
      commandId: command.commandId,
      operation: 'apply',
      documentEpoch: this.documentEpoch,
      beforeChangeSeq: beforeSeq,
      afterChangeSeq: this.changeSeq,
      beforeDocumentSha256: state.documentSha256,
      afterDocumentSha256,
      beforeTextSha256: beforeEvidence.textSha256,
      afterTextSha256: afterEvidence.textSha256,
      formatSha256: afterEvidence.formatSha256,
      adjacentContextSha256: afterEvidence.adjacentContextSha256,
      pageCountBefore: state.pageCount,
      pageCountAfter: this.deps.wasm.pageCount,
      target: command.target,
    };
    const entry: FieldAgentJournalEntry = {
      command: structuredClone(command),
      applyBindingSha256: binding,
      applyReceipt: receipt,
      beforeText: beforeEvidence.text,
      beforeEvidence,
      afterEvidence,
      nonTargetManifestSha256: beforeNonTarget,
      status: 'applied',
    };
    this.fieldJournal.set(command.commandId, entry);
    this.latestField = entry;
    this.emitDocumentAgentChanged({
      schemaVersion: 1,
      reason: 'field_agent_apply',
      documentEpoch: this.documentEpoch,
      changeSeq: this.changeSeq,
      commandId: command.commandId,
    });
    return receipt;
  }

  async revertFieldCommand(command: RhwpRevertFieldCommandV1): Promise<RhwpFieldCommandReceiptV1> {
    this.syncGeneration();
    const binding = revertBinding(command);
    const entry = this.fieldJournal.get(command.commandId);
    if (entry?.status === 'reverted') {
      if (entry.revertBindingSha256 === binding
          && entry.revertReceipt
          && this.isCurrentFieldReceiptState(
            entry.revertReceipt,
            entry.revertReceipt.afterTextSha256,
          )) return entry.revertReceipt;
      throw new DocumentAgentError('COMMAND_REPLAY_MISMATCH', 'field revert replay binding이 다릅니다.');
    }
    if (!entry || this.latestField !== entry || entry.status !== 'applied'
        || this.changeSeq !== entry.applyReceipt.afterChangeSeq) {
      throw new DocumentAgentError('COMMAND_NOT_LATEST', '가장 최근 exact field command만 되돌릴 수 있습니다.');
    }

    const startedAt = this.now();
    const state = this.getDocumentState();
    if (command.expectedDocumentEpoch !== state.documentEpoch) {
      throw new DocumentAgentError('DOCUMENT_EPOCH_MISMATCH', 'document epoch가 다릅니다.');
    }
    if (command.expectedChangeSeq !== state.changeSeq) {
      throw new DocumentAgentError('CHANGE_SEQ_MISMATCH', 'changeSeq가 다릅니다.');
    }
    if (command.expectedAfterDocumentSha256 !== state.documentSha256) {
      throw new DocumentAgentError('DOCUMENT_SHA_MISMATCH', 'field after document SHA가 다릅니다.');
    }
    const currentEvidence = collectFieldTargetEvidence(this.deps.wasm, entry.command.target);
    if (currentEvidence.textSha256 !== command.expectedAfterSha256
        || currentEvidence.textSha256 !== entry.applyReceipt.afterTextSha256
        || currentEvidence.adjacentContextSha256 !== entry.nonTargetManifestSha256) {
      throw new DocumentAgentError('COMMAND_NOT_LATEST', 'field apply 뒤 문서가 변경되어 되돌릴 수 없습니다.');
    }

    let revertedEvidence!: TargetEvidence;
    let afterDocumentSha256 = '';
    const beforeSeq = this.changeSeq;
    try {
      await this.deps.input.executeDocumentAgentOperation({
        kind: 'snapshot',
        operationType: 'field-agent:revert',
        operation: (wasm) => {
          this.assertRuntimeFence(command.expectedDocumentEpoch, beforeSeq);
          let deferred = false;
          try {
            wasm.beginDeferredPagination?.();
            deferred = true;
            const result = replaceWholeFieldTextDeferred(
              wasm,
              entry.command.target,
              currentEvidence.text,
              entry.beforeText,
            );
            if (!result.ok || result.charOffset !== codePointLength(entry.beforeText)) {
              throw new DocumentAgentError('TRANSACTION_FAILED', 'field inverse replace가 실패했습니다.');
            }
            const beforeLength = codePointLength(entry.beforeText);
            if (beforeLength > 0) {
              wasm.setCharShapeIdInCell(
                entry.command.target.section,
                entry.command.target.parentPara,
                entry.command.target.controlIndex,
                entry.command.target.cellIndex,
                entry.command.target.cellParagraph,
                0,
                beforeLength,
                entry.beforeEvidence.charShapeId,
              );
            }
            wasm.setCellParaShapeId(
              entry.command.target.section,
              entry.command.target.parentPara,
              entry.command.target.controlIndex,
              entry.command.target.cellIndex,
              entry.command.target.cellParagraph,
              entry.beforeEvidence.paraShapeId,
            );
            wasm.flushDeferredPagination?.();
            deferred = false;
            revertedEvidence = collectFieldTargetEvidence(wasm, entry.command.target);
            if (revertedEvidence.textSha256 !== entry.beforeEvidence.textSha256
                || revertedEvidence.formatSha256 !== entry.beforeEvidence.formatSha256
                || revertedEvidence.adjacentContextSha256 !== entry.nonTargetManifestSha256) {
              throw new DocumentAgentError('TRANSACTION_FAILED', 'field before 상태 복원이 일치하지 않습니다.');
            }
            if (wasm.pageCount !== entry.applyReceipt.pageCountBefore) {
              throw new DocumentAgentError('PAGE_COUNT_CHANGED', 'field revert 뒤 페이지 수가 다릅니다.');
            }
            afterDocumentSha256 = exportDocumentSha256(wasm, state.format);
            this.assertWithinBudget(startedAt);
            return {
              sectionIndex: entry.command.target.section,
              paragraphIndex: entry.command.target.cellParagraph,
              charOffset: beforeLength,
              parentParaIndex: entry.command.target.parentPara,
              controlIndex: entry.command.target.controlIndex,
              cellIndex: entry.command.target.cellIndex,
              cellParaIndex: entry.command.target.cellParagraph,
            };
          } catch (error) {
            if (deferred) {
              try { wasm.cancelDeferredPagination?.(); } catch { /* snapshot rollback이 최종 복구한다. */ }
            }
            throw error;
          }
        },
        meta: {
          actionId: 'field-agent:revert',
          domain: 'text',
          refresh: 'full',
          dirtyScope: 'paragraph',
          selection: 'moveToResult',
        },
      }, async () => {
        await this.deps.render();
        this.assertWithinBudget(startedAt);
      });
    } catch (error) {
      throw this.normalizeExecutionError(error);
    }
    if (this.changeSeq !== beforeSeq + 1) {
      throw new DocumentAgentError('TRANSACTION_FAILED', 'field revert changeSeq가 정확히 1 증가하지 않았습니다.', false);
    }
    afterDocumentSha256 = exportPersistedDocumentSha256(this.deps.wasm, state.format);

    const receipt: RhwpFieldCommandReceiptV1 = {
      schemaVersion: 1,
      commandId: command.commandId,
      operation: 'revert',
      documentEpoch: this.documentEpoch,
      beforeChangeSeq: beforeSeq,
      afterChangeSeq: this.changeSeq,
      beforeDocumentSha256: state.documentSha256,
      afterDocumentSha256,
      beforeTextSha256: currentEvidence.textSha256,
      afterTextSha256: revertedEvidence.textSha256,
      formatSha256: revertedEvidence.formatSha256,
      adjacentContextSha256: revertedEvidence.adjacentContextSha256,
      pageCountBefore: state.pageCount,
      pageCountAfter: this.deps.wasm.pageCount,
      target: entry.command.target,
    };
    entry.status = 'reverted';
    entry.revertBindingSha256 = binding;
    entry.revertReceipt = receipt;
    this.emitDocumentAgentChanged({
      schemaVersion: 1,
      reason: 'field_agent_revert',
      documentEpoch: this.documentEpoch,
      changeSeq: this.changeSeq,
      commandId: command.commandId,
    });
    return receipt;
  }

  focusTarget(target: RhwpBodyParagraphTargetV1): { focused: boolean; page: number } {
    this.syncGeneration();
    assertTargetCoordinates(this.deps.wasm, target);
    const page = this.pageFor(target.section, target.paragraph);
    return {
      focused: this.deps.input.focusBodyParagraph(target.section, target.paragraph, target.length),
      page,
    };
  }

  focusFieldTarget(target: RhwpTableCellTextTargetV1): { focused: boolean; page: number } {
    this.syncGeneration();
    return this.deps.input.focusTableCellText(
      target.section,
      target.parentPara,
      target.controlIndex,
      target.cellIndex,
      target.cellParagraph,
    );
  }

  private currentFormat(): 'hwp' | 'hwpx' {
    const format = this.deps.wasm.getSourceFormat();
    if (format !== 'hwp' && format !== 'hwpx') {
      throw new DocumentAgentError(
        'CAPABILITY_UNSUPPORTED',
        `document agent는 HWP/HWPX만 지원합니다: ${format}`,
      );
    }
    return format;
  }

  private pageFor(section: number, paragraph: number): number {
    const result = this.deps.wasm.getPageOfPosition(section, paragraph);
    if (!result.ok || !Number.isSafeInteger(result.page) || (result.page as number) < 0) {
      return 1;
    }
    return (result.page as number) + 1;
  }

  private assertApplyFence(
    command: Pick<RhwpApplyTextCommandV1, 'expectedDocumentEpoch' | 'expectedChangeSeq' | 'expectedDocumentSha256'>,
    state: RhwpDocumentStateV1,
  ): void {
    if (command.expectedDocumentEpoch !== state.documentEpoch) {
      throw new DocumentAgentError('DOCUMENT_EPOCH_MISMATCH', 'document epoch가 다릅니다.');
    }
    if (command.expectedChangeSeq !== state.changeSeq) {
      throw new DocumentAgentError('CHANGE_SEQ_MISMATCH', 'changeSeq가 다릅니다.');
    }
    if (command.expectedDocumentSha256 !== state.documentSha256) {
      throw new DocumentAgentError('DOCUMENT_SHA_MISMATCH', 'document SHA가 다릅니다.');
    }
  }

  private assertRuntimeFence(expectedEpoch: number, expectedSeq: number): void {
    this.syncGeneration();
    if (this.documentEpoch !== expectedEpoch) {
      throw new DocumentAgentError('DOCUMENT_EPOCH_MISMATCH', 'transaction 직전 epoch가 바뀌었습니다.');
    }
    if (this.changeSeq !== expectedSeq) {
      throw new DocumentAgentError('CHANGE_SEQ_MISMATCH', 'transaction 직전 changeSeq가 바뀌었습니다.');
    }
  }

  private assertWithinBudget(startedAt: number): void {
    if (this.now() - startedAt > COMMAND_BUDGET_MS) {
      throw new DocumentAgentError('COMMAND_TOO_SLOW', '문서 명령이 3초 상한을 넘었습니다.');
    }
  }

  private syncGeneration(): void {
    const generation = this.deps.wasm.documentGeneration;
    if (generation === this.documentEpoch) return;
    this.documentEpoch = generation;
    this.changeSeq = 0;
    this.latest = null;
    this.journal.clear();
    this.latestField = null;
    this.fieldJournal.clear();
  }

  private normalizeExecutionError(error: unknown): DocumentAgentError {
    if (isDocumentAgentError(error)) {
      if (error.code === 'TRANSACTION_FAILED' && error.recovered === undefined) {
        return new DocumentAgentError(error.code, error.message, true);
      }
      return error;
    }
    const executionError = error as {
      code?: unknown;
      recovered?: unknown;
      cause?: unknown;
      message?: unknown;
    };
    const nestedCause = executionError.cause as { code?: unknown; message?: unknown } | undefined;
    if (executionError.code === 'RENDER_FAILED'
        && typeof executionError.recovered === 'boolean') {
      if (nestedCause?.code === 'COMMAND_TOO_SLOW') {
        return new DocumentAgentError(
          'COMMAND_TOO_SLOW',
          typeof nestedCause.message === 'string' ? nestedCause.message : 'command time budget을 초과했습니다.',
          executionError.recovered,
        );
      }
      return new DocumentAgentError(
        'RENDER_FAILED',
        typeof executionError.message === 'string' ? executionError.message : 'render commit이 실패했습니다.',
        executionError.recovered,
      );
    }
    return new DocumentAgentError(
      'TRANSACTION_FAILED',
      error instanceof Error ? error.message : String(error),
      !(error instanceof AggregateError),
    );
  }

  private isCurrentReceiptState(
    receipt: RhwpTextCommandReceiptV1,
    target: RhwpBodyParagraphTargetV1,
    expectedTextSha256: string,
  ): boolean {
    try {
      const state = this.getDocumentState();
      if (state.documentEpoch !== receipt.documentEpoch
          || state.changeSeq !== receipt.afterChangeSeq
          || state.documentSha256 !== receipt.afterDocumentSha256) return false;
      return collectTargetEvidence(this.deps.wasm, target).textSha256 === expectedTextSha256;
    } catch {
      return false;
    }
  }

  private isCurrentFieldReceiptState(
    receipt: RhwpFieldCommandReceiptV1,
    expectedTextSha256: string,
  ): boolean {
    try {
      const state = this.getDocumentState();
      if (state.documentEpoch !== receipt.documentEpoch
          || state.changeSeq !== receipt.afterChangeSeq
          || state.documentSha256 !== receipt.afterDocumentSha256) return false;
      return collectFieldTargetEvidence(this.deps.wasm, receipt.target).textSha256 === expectedTextSha256;
    } catch {
      return false;
    }
  }

  private emitDocumentAgentChanged(event: {
    schemaVersion: 1;
    reason: 'agent_apply' | 'agent_revert' | 'field_agent_apply' | 'field_agent_revert';
    documentEpoch: number;
    changeSeq: number;
    commandId: string;
  }): void {
    try {
      this.deps.eventBus.emit('document-agent-changed', event);
    } catch (error) {
      // transaction과 journal은 이미 commit됐다. 관측자 한 곳의 실패로 RPC 성공을 뒤집지 않는다.
      console.error('[DocumentAgentController] documentChanged observer 실패:', error);
    }
  }
}
