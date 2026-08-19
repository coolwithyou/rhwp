import { resolve } from 'node:path';

import { assert, runTest } from './helpers.mjs';

const LOCAL_URL = process.env.VITE_URL || 'http://localhost:7700';
const STUDIO_URL = process.env.STUDIO_URL;
if (!STUDIO_URL) throw new Error('STUDIO_URL is required');

const editorModulePath = resolve(import.meta.dirname, '../../npm/editor/index.js')
  .replace(/\\/g, '/');
const editorModuleUrl = `${LOCAL_URL}/@fs${editorModulePath}`;
const controllerModuleUrl = `${LOCAL_URL}/src/document-agent/controller.ts`;

await runTest('production Studio form_text public SDK apply/reopen/revert smoke', async ({ page }) => {
  page.on('console', message => console.log(`  [production browser] ${message.text()}`));
  await page.goto(`${LOCAL_URL}/e2e/embed-harness.html`, { waitUntil: 'domcontentloaded' });

  const result = await page.evaluate(async ({
    editorModuleUrl: sdkUrl,
    controllerModuleUrl: evidenceUrl,
    localUrl,
    studioUrl,
  }) => {
    const { createEditor } = await import(sdkUrl);
    const { collectFieldTargetEvidence } = await import(evidenceUrl);
    const sourceBytes = await fetch(`${localUrl}/samples/field-01.hwp`)
      .then(response => response.arrayBuffer());

    const createHost = (hidden = false) => {
      const host = document.createElement('div');
      host.style.cssText = hidden
        ? 'position:fixed;width:1024px;height:768px;overflow:hidden;left:-2000px;top:0'
        : 'width:100vw;height:100vh';
      document.body.append(host);
      return host;
    };

    document.body.replaceChildren();
    const evidenceHost = createHost(true);
    const evidenceEditor = await createEditor(evidenceHost, {
      studioUrl: `${localUrl.replace(/\/$/, '')}/`,
      renderer: 'canvas2d',
      handshakeTimeoutMs: 10_000,
    });
    await evidenceEditor.loadFile(sourceBytes, 'field-01.hwp', { suppressDialogs: true });
    const evidenceWasm = evidenceEditor.element.contentWindow.__wasm;
    if (!evidenceWasm) throw new Error('local evidence WasmBridge is unavailable');
    const field = evidenceWasm.getFieldList().find(entry =>
      entry.name === '회사명'
      && entry.fieldType === 'clickhere'
      && entry.editableInForm === true
      && !(entry.location.path?.length));
    if (!field) throw new Error('safe root form_text field not found');
    const target = {
      kind: 'form_text',
      section: field.location.sectionIndex,
      paragraph: field.location.paraIndex,
      fieldId: field.fieldId,
    };
    const before = collectFieldTargetEvidence(evidenceWasm, target);
    const evidenceState = await evidenceEditor.getDocumentState();
    evidenceEditor.destroy();
    evidenceHost.remove();

    const productionHost = createHost();
    const editor = await createEditor(productionHost, {
      studioUrl: `${studioUrl.replace(/\/$/, '')}/`,
      renderer: 'canvas2d',
      handshakeTimeoutMs: 15_000,
    });
    await editor.loadFile(sourceBytes, 'field-01.hwp', { suppressDialogs: true });
    const initialState = await editor.getDocumentState();
    const documentEvents = [];
    const selectionEvents = [];
    const offDocument = editor.onDocumentChanged(event => documentEvents.push(event.reason));
    const offSelection = editor.onFieldSelectionChanged(event => selectionEvents.push(event.target));
    const focus = await editor.focusFieldTarget(target);
    await new Promise(resolveFrame => requestAnimationFrame(resolveFrame));
    const focusedSelection = await editor.getFieldSelectionContext();
    const commandId = crypto.randomUUID();
    console.log('initial apply start');
    const applied = await editor.applyFieldCommand({
      schemaVersion: 1,
      commandId,
      expectedDocumentEpoch: initialState.documentEpoch,
      expectedChangeSeq: initialState.changeSeq,
      expectedDocumentSha256: initialState.documentSha256,
      target,
      expectedBeforeSha256: before.textSha256,
      expectedFormatSha256: before.formatSha256,
      expectedAdjacentContextSha256: before.adjacentContextSha256,
      replacement: '주식회사 노튼',
    });
    console.log('initial apply complete');
    const appliedBytes = await editor.exportHwp();

    const reopenHost = createHost(true);
    const reopened = await createEditor(reopenHost, {
      studioUrl: `${studioUrl.replace(/\/$/, '')}/`,
      renderer: 'canvas2d',
      handshakeTimeoutMs: 15_000,
    });
    await reopened.loadFile(appliedBytes, 'field-01-applied.hwp', { suppressDialogs: true });
    const reopenedState = await reopened.getDocumentState();
    const reopenCommandId = crypto.randomUUID();
    console.log('reopen probe apply start');
    const reopenedApplied = await reopened.applyFieldCommand({
      schemaVersion: 1,
      commandId: reopenCommandId,
      expectedDocumentEpoch: reopenedState.documentEpoch,
      expectedChangeSeq: reopenedState.changeSeq,
      expectedDocumentSha256: reopenedState.documentSha256,
      target,
      expectedBeforeSha256: applied.afterTextSha256,
      expectedFormatSha256: applied.formatSha256,
      expectedAdjacentContextSha256: applied.adjacentContextSha256,
      replacement: '노튼 주식회사',
    });
    console.log('reopen probe apply complete');
    await reopened.revertFieldCommand({
      schemaVersion: 1,
      commandId: reopenCommandId,
      expectedDocumentEpoch: reopenedApplied.documentEpoch,
      expectedChangeSeq: reopenedApplied.afterChangeSeq,
      expectedAfterDocumentSha256: reopenedApplied.afterDocumentSha256,
      expectedAfterSha256: reopenedApplied.afterTextSha256,
    });
    console.log('reopen probe revert complete');
    reopened.destroy();
    reopenHost.remove();

    console.log('initial revert start');
    const reverted = await editor.revertFieldCommand({
      schemaVersion: 1,
      commandId,
      expectedDocumentEpoch: applied.documentEpoch,
      expectedChangeSeq: applied.afterChangeSeq,
      expectedAfterDocumentSha256: applied.afterDocumentSha256,
      expectedAfterSha256: applied.afterTextSha256,
    });
    console.log('initial revert complete');
    const restoredBytes = await editor.exportHwp();

    const restoredHost = createHost(true);
    const restored = await createEditor(restoredHost, {
      studioUrl: `${studioUrl.replace(/\/$/, '')}/`,
      renderer: 'canvas2d',
      handshakeTimeoutMs: 15_000,
    });
    await restored.loadFile(restoredBytes, 'field-01-restored.hwp', { suppressDialogs: true });
    const restoredState = await restored.getDocumentState();
    const restoreProbeId = crypto.randomUUID();
    console.log('restore probe apply start');
    const restoreProbe = await restored.applyFieldCommand({
      schemaVersion: 1,
      commandId: restoreProbeId,
      expectedDocumentEpoch: restoredState.documentEpoch,
      expectedChangeSeq: restoredState.changeSeq,
      expectedDocumentSha256: restoredState.documentSha256,
      target,
      expectedBeforeSha256: before.textSha256,
      expectedFormatSha256: before.formatSha256,
      expectedAdjacentContextSha256: before.adjacentContextSha256,
      replacement: '주식회사 노튼',
    });
    console.log('restore probe apply complete');
    await restored.revertFieldCommand({
      schemaVersion: 1,
      commandId: restoreProbeId,
      expectedDocumentEpoch: restoreProbe.documentEpoch,
      expectedChangeSeq: restoreProbe.afterChangeSeq,
      expectedAfterDocumentSha256: restoreProbe.afterDocumentSha256,
      expectedAfterSha256: restoreProbe.afterTextSha256,
    });
    console.log('restore probe revert complete');
    restored.destroy();
    restoredHost.remove();

    offDocument();
    offSelection();
    editor.destroy();
    productionHost.remove();

    return {
      sourceShaMatched: initialState.documentSha256 === evidenceState.documentSha256,
      focus,
      focusedSelection,
      target,
      applied,
      reopenedApplied,
      reverted,
      restoreProbe,
      documentEvents,
      selectionEventMatched: selectionEvents.some(selection =>
        selection?.kind === 'form_text' && selection.fieldId === target.fieldId),
    };
  }, {
    editorModuleUrl,
    controllerModuleUrl,
    localUrl: LOCAL_URL,
    studioUrl: STUDIO_URL,
  });

  assert(result.sourceShaMatched, '로컬 증거와 프로덕션 Studio source SHA 일치');
  assert(result.focus.focused, '프로덕션 Studio exact form_text focus 성공');
  assert(result.focusedSelection.editable
    && result.focusedSelection.target?.fieldId === result.target.fieldId,
  '프로덕션 Studio selection context가 exact form_text를 반환');
  assert(result.applied.operation === 'apply'
    && result.applied.afterChangeSeq === result.applied.beforeChangeSeq + 1,
  '프로덕션 Studio form_text apply receipt 정상');
  assert(result.reopenedApplied.beforeTextSha256 === result.applied.afterTextSha256,
    '프로덕션 적용본 HWP 재개방 뒤 exact 값 증거 유지');
  assert(result.reverted.operation === 'revert'
    && result.reverted.afterTextSha256 === result.applied.beforeTextSha256,
  '프로덕션 Studio form_text exact revert 정상');
  assert(result.restoreProbe.beforeTextSha256 === result.applied.beforeTextSha256,
    '프로덕션 복원본 HWP 재개방 뒤 원문 증거 유지');
  assert(JSON.stringify(result.documentEvents) === JSON.stringify([
    'field_agent_apply', 'field_agent_revert',
  ]), '프로덕션 Studio public document event 순서');
  assert(result.selectionEventMatched,
    '프로덕션 Studio selection event에 exact form_text target 전달');
}, { skipLoadApp: true });
