import test from 'node:test';
import assert from 'node:assert/strict';

import { RhwpEditor } from '../index.js';

const SHA_A = 'a'.repeat(64);
const SHA_B = 'b'.repeat(64);
const SHA_C = 'c'.repeat(64);
const SHA_D = 'd'.repeat(64);

function target() {
  return {
    kind: 'body_paragraph',
    section: 0,
    paragraph: 2,
    charOffset: 0,
    length: 7,
  };
}

function fieldTarget() {
  return {
    kind: 'table_cell_text',
    section: 0,
    parentPara: 4,
    controlIndex: 1,
    cellIndex: 3,
    cellParagraph: 0,
  };
}

function formFieldTarget() {
  return {
    kind: 'form_text',
    section: 0,
    paragraph: 5,
    fieldId: 17,
  };
}

function fieldRegionTarget() {
  return {
    kind: 'table_cell_region',
    section: 0,
    parentPara: 4,
    controlIndex: 1,
    cellIndex: 3,
  };
}

function state() {
  return {
    schemaVersion: 1,
    format: 'hwp',
    documentEpoch: 3,
    changeSeq: 11,
    dirty: true,
    pageCount: 2,
    documentSha256: SHA_A,
  };
}

function selection() {
  return {
    schemaVersion: 1,
    documentEpoch: 3,
    changeSeq: 11,
    page: 2,
    editable: true,
    collapsed: true,
    target: target(),
    selectedTextSha256: null,
  };
}

function fieldSelection(selectedTarget = fieldTarget()) {
  return {
    schemaVersion: 1,
    documentEpoch: 3,
    changeSeq: 11,
    page: 2,
    editable: selectedTarget !== null,
    target: selectedTarget,
  };
}

function applyCommand() {
  return {
    schemaVersion: 1,
    commandId: 'cmd-001',
    expectedDocumentEpoch: 3,
    expectedChangeSeq: 11,
    expectedDocumentSha256: SHA_A,
    target: target(),
    expectedBeforeSha256: SHA_B,
    expectedFormatSha256: SHA_C,
    expectedAdjacentContextSha256: SHA_D,
    replacement: '새 문단',
  };
}

function revertCommand() {
  return {
    schemaVersion: 1,
    commandId: 'cmd-001',
    expectedDocumentEpoch: 3,
    expectedChangeSeq: 12,
    expectedAfterDocumentSha256: SHA_B,
    expectedAfterSha256: SHA_C,
  };
}

function applyFieldCommand() {
  return { ...applyCommand(), commandId: 'field-001', target: fieldTarget(), replacement: '주식회사 노튼' };
}

function applyFormFieldCommand() {
  return {
    ...applyCommand(),
    commandId: 'form-field-001',
    target: formFieldTarget(),
    replacement: '주식회사 노튼',
  };
}

function revertFieldCommand() {
  return { ...revertCommand(), commandId: 'field-001' };
}

function receipt(operation = 'apply') {
  return {
    schemaVersion: 1,
    commandId: 'cmd-001',
    operation,
    documentEpoch: 3,
    beforeChangeSeq: operation === 'apply' ? 11 : 12,
    afterChangeSeq: operation === 'apply' ? 12 : 13,
    beforeDocumentSha256: SHA_A,
    afterDocumentSha256: SHA_B,
    beforeTextSha256: SHA_C,
    afterTextSha256: SHA_D,
    formatSha256: SHA_A,
    adjacentContextSha256: SHA_B,
    pageCountBefore: 2,
    pageCountAfter: 2,
    target: target(),
  };
}

function fieldReceipt(operation = 'apply') {
  return { ...receipt(operation), commandId: 'field-001', target: fieldTarget() };
}

function formFieldReceipt(operation = 'apply') {
  return {
    ...receipt(operation),
    commandId: 'form-field-001',
    target: formFieldTarget(),
  };
}

function editorHarness(results, capabilities = [
  'document-state-v1',
  'selection-context-v1',
  'document-agent-command-v1',
  'target-navigation-v1',
  'field-target-navigation-v1',
  'field-agent-command-v1',
  'field-selection-events-v1',
  'document-change-events-v1',
]) {
  const requests = [];
  const listeners = new Map();
  const transport = {
    request(method, params) {
      requests.push({ method, params });
      return Promise.resolve(results[method]);
    },
    supports(capability) { return capabilities.includes(capability); },
    on(event, listener) {
      listeners.set(event, listener);
      return () => listeners.delete(event);
    },
    destroy() {},
  };
  return {
    editor: new RhwpEditor({ remove() {} }, transport),
    requests,
    emit(event, payload) { listeners.get(event)?.(payload); },
    listeners,
  };
}

test('문서 에이전트 공개 API는 exact RPC 메서드와 파라미터를 사용한다', async () => {
  const results = {
    getDocumentState: state(),
    getSelectionContext: selection(),
    getFieldSelectionContext: fieldSelection(),
    applyTextCommand: receipt('apply'),
    revertTextCommand: receipt('revert'),
    focusTarget: { focused: true, page: 2 },
    focusFieldTarget: { focused: true, page: 3 },
    applyFieldCommand: fieldReceipt('apply'),
    revertFieldCommand: fieldReceipt('revert'),
  };
  const { editor, requests } = editorHarness(results);

  assert.deepEqual(await editor.getDocumentState(), state());
  assert.deepEqual(await editor.getSelectionContext(), selection());
  assert.deepEqual(await editor.getFieldSelectionContext(), fieldSelection());
  assert.deepEqual(await editor.applyTextCommand(applyCommand()), receipt('apply'));
  assert.deepEqual(await editor.revertTextCommand(revertCommand()), receipt('revert'));
  assert.deepEqual(await editor.focusTarget(target()), { focused: true, page: 2 });
  assert.deepEqual(await editor.focusFieldTarget(fieldTarget()), { focused: true, page: 3 });
  assert.deepEqual(await editor.applyFieldCommand(applyFieldCommand()), fieldReceipt('apply'));
  assert.deepEqual(await editor.revertFieldCommand(revertFieldCommand()), fieldReceipt('revert'));

  assert.deepEqual(requests, [
    { method: 'getDocumentState', params: {} },
    { method: 'getSelectionContext', params: {} },
    { method: 'getFieldSelectionContext', params: {} },
    { method: 'applyTextCommand', params: { command: applyCommand() } },
    { method: 'revertTextCommand', params: { command: revertCommand() } },
    { method: 'focusTarget', params: { target: target() } },
    { method: 'focusFieldTarget', params: { target: fieldTarget() } },
    { method: 'applyFieldCommand', params: { command: applyFieldCommand() } },
    { method: 'revertFieldCommand', params: { command: revertFieldCommand() } },
  ]);
});

test('문서 에이전트 공개 API는 exact 본문 누름틀 target을 전달하고 검증한다', async () => {
  const results = {
    getFieldSelectionContext: fieldSelection(formFieldTarget()),
    focusFieldTarget: { focused: true, page: 4 },
    applyFieldCommand: formFieldReceipt('apply'),
  };
  const { editor, requests } = editorHarness(results);

  assert.deepEqual(await editor.getFieldSelectionContext(), fieldSelection(formFieldTarget()));
  assert.deepEqual(await editor.focusFieldTarget(formFieldTarget()), { focused: true, page: 4 });
  assert.deepEqual(await editor.applyFieldCommand(applyFormFieldCommand()), formFieldReceipt('apply'));
  assert.deepEqual(requests, [
    { method: 'getFieldSelectionContext', params: {} },
    { method: 'focusFieldTarget', params: { target: formFieldTarget() } },
    { method: 'applyFieldCommand', params: { command: applyFormFieldCommand() } },
  ]);
});

test('문서 에이전트 공개 API는 표 셀 장문 target의 줄바꿈을 허용한다', async () => {
  const command = {
    ...applyFieldCommand(),
    commandId: 'field-region-001',
    target: fieldRegionTarget(),
    replacement: '첫 문단\n둘째 문단',
  };
  const result = {
    ...fieldReceipt('apply'),
    commandId: command.commandId,
    target: fieldRegionTarget(),
  };
  const { editor, requests } = editorHarness({
    focusFieldTarget: { focused: true, page: 3 },
    applyFieldCommand: result,
  });

  assert.deepEqual(await editor.focusFieldTarget(fieldRegionTarget()), { focused: true, page: 3 });
  assert.deepEqual(await editor.applyFieldCommand(command), result);
  assert.deepEqual(requests, [
    { method: 'focusFieldTarget', params: { target: fieldRegionTarget() } },
    { method: 'applyFieldCommand', params: { command } },
  ]);
});

test('문서 에이전트 공개 API는 capability가 없으면 요청 전에 실패한다', async () => {
  const { editor, requests } = editorHarness({}, []);

  for (const call of [
    () => editor.getDocumentState(),
    () => editor.getSelectionContext(),
    () => editor.getFieldSelectionContext(),
    () => editor.applyTextCommand(applyCommand()),
    () => editor.revertTextCommand(revertCommand()),
    () => editor.focusTarget(target()),
    () => editor.focusFieldTarget(fieldTarget()),
    () => editor.applyFieldCommand(applyFieldCommand()),
    () => editor.revertFieldCommand(revertFieldCommand()),
  ]) {
    await assert.rejects(call, (error) => error.code === 'CAPABILITY_UNSUPPORTED');
  }
  assert.deepEqual(requests, []);
});

test('문서 에이전트 공개 API는 extra key와 잘못된 SHA를 요청 전에 거부한다', async () => {
  const { editor, requests } = editorHarness({});

  await assert.rejects(
    () => editor.applyTextCommand({ ...applyCommand(), unexpected: true }),
    (error) => error.code === 'INVALID_COMMAND',
  );
  await assert.rejects(
    () => editor.applyTextCommand({ ...applyCommand(), expectedDocumentSha256: 'not-a-sha' }),
    (error) => error.code === 'INVALID_COMMAND',
  );
  await assert.rejects(
    () => editor.applyTextCommand({ ...applyCommand(), replacement: '두 문단\n금지' }),
    (error) => error.code === 'INVALID_COMMAND',
  );
  await assert.rejects(
    () => editor.focusTarget({ ...target(), charOffset: 1 }),
    (error) => error.code === 'INVALID_COMMAND',
  );
  await assert.rejects(
    () => editor.focusFieldTarget({ ...fieldTarget(), cellIndex: -1 }),
    (error) => error.code === 'INVALID_COMMAND',
  );
  await assert.rejects(
    () => editor.applyFieldCommand({ ...applyFieldCommand(), replacement: '두 줄\n금지' }),
    (error) => error.code === 'INVALID_COMMAND',
  );
  await assert.rejects(
    () => editor.applyFieldCommand({ ...applyFieldCommand(), target: { ...fieldTarget(), page: 1 } }),
    (error) => error.code === 'INVALID_COMMAND',
  );
  await assert.rejects(
    () => editor.focusFieldTarget({ ...formFieldTarget(), fieldId: -1 }),
    (error) => error.code === 'INVALID_COMMAND',
  );
  await assert.rejects(
    () => editor.applyFieldCommand({ ...applyFormFieldCommand(), target: { ...formFieldTarget(), extra: true } }),
    (error) => error.code === 'INVALID_COMMAND',
  );
  assert.deepEqual(requests, []);
});

test('문서 에이전트 공개 API는 malformed 응답을 명시적으로 거부한다', async () => {
  const malformedState = { ...state(), pageCount: Number.NaN };
  const malformedSelection = { ...selection(), page: 0 };
  const malformedFieldSelection = { ...fieldSelection(), unexpected: true };
  const malformedReceipt = { ...receipt(), extra: true };
  const malformedFocus = { focused: true, page: 1, extra: true };
  const { editor } = editorHarness({
    getDocumentState: malformedState,
    getSelectionContext: malformedSelection,
    getFieldSelectionContext: malformedFieldSelection,
    applyTextCommand: malformedReceipt,
    focusTarget: malformedFocus,
    focusFieldTarget: malformedFocus,
    applyFieldCommand: malformedReceipt,
  });

  await assert.rejects(
    () => editor.getDocumentState(),
    (error) => error.code === 'INVALID_RESPONSE',
  );
  await assert.rejects(
    () => editor.getSelectionContext(),
    (error) => error.code === 'INVALID_RESPONSE',
  );
  await assert.rejects(
    () => editor.getFieldSelectionContext(),
    (error) => error.code === 'INVALID_RESPONSE',
  );
  await assert.rejects(
    () => editor.applyTextCommand(applyCommand()),
    (error) => error.code === 'INVALID_RESPONSE',
  );
  await assert.rejects(
    () => editor.focusTarget(target()),
    (error) => error.code === 'INVALID_RESPONSE',
  );
  await assert.rejects(
    () => editor.focusFieldTarget(fieldTarget()),
    (error) => error.code === 'INVALID_RESPONSE',
  );
  await assert.rejects(
    () => editor.applyFieldCommand(applyFieldCommand()),
    (error) => error.code === 'INVALID_RESPONSE',
  );
});

test('문서 변경 이벤트는 capability와 strict v1 payload를 사용한다', () => {
  const { editor, emit, listeners } = editorHarness({});
  const received = [];
  const off = editor.onDocumentChanged((event) => received.push(event));
  const event = {
    schemaVersion: 1,
    reason: 'agent_apply',
    documentEpoch: 3,
    changeSeq: 12,
    commandId: 'cmd-001',
  };
  emit('documentChanged', event);
  assert.deepEqual(received, [event]);
  const fieldEvent = { ...event, reason: 'field_agent_apply', changeSeq: 13, commandId: 'field-001' };
  emit('documentChanged', fieldEvent);
  assert.deepEqual(received, [event, fieldEvent]);
  off();
  assert.equal(listeners.has('documentChanged'), false);

  const unsupported = editorHarness({}, []).editor;
  assert.throws(
    () => unsupported.onDocumentChanged(() => {}),
    (error) => error.code === 'CAPABILITY_UNSUPPORTED',
  );
});

test('문서 변경 이벤트는 모든 listener에 한 번 전달하고 stale epoch/seq를 버린다', () => {
  const { editor, emit } = editorHarness({});
  const first = [];
  const second = [];
  editor.onDocumentChanged(event => first.push(event.changeSeq));
  editor.onDocumentChanged(event => second.push(event.changeSeq));

  const event = (documentEpoch, changeSeq) => ({
    schemaVersion: 1,
    reason: 'agent_apply',
    documentEpoch,
    changeSeq,
    commandId: `cmd-${documentEpoch}-${changeSeq}`,
  });
  emit('documentChanged', event(3, 2));
  emit('documentChanged', event(3, 2));
  emit('documentChanged', event(2, 99));
  emit('documentChanged', event(3, 3));

  assert.deepEqual(first, [2, 3]);
  assert.deepEqual(second, [2, 3]);
});

test('field selection 이벤트는 같은 revision의 서로 다른 셀 이동도 모두 전달한다', () => {
  const { editor, emit, listeners } = editorHarness({});
  const received = [];
  const off = editor.onFieldSelectionChanged(
    event => received.push(event.target?.cellIndex ?? null),
  );

  emit('fieldSelectionChanged', fieldSelection({ ...fieldTarget(), cellIndex: 3 }));
  emit('fieldSelectionChanged', fieldSelection({ ...fieldTarget(), cellIndex: 5 }));
  emit('fieldSelectionChanged', fieldSelection(null));
  emit('fieldSelectionChanged', { ...fieldSelection(), unexpected: true });

  assert.deepEqual(received, [3, 5, null]);
  off();
  assert.equal(listeners.has('fieldSelectionChanged'), false);

  const unsupported = editorHarness({}, []).editor;
  assert.throws(
    () => unsupported.onFieldSelectionChanged(() => {}),
    (error) => error.code === 'CAPABILITY_UNSUPPORTED',
  );
});
