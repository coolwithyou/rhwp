import { resolve } from 'node:path';

import { runTest, assert } from './helpers.mjs';

const EDITOR_MODULE_PATH = resolve(import.meta.dirname, '../../npm/editor/index.js').replace(/\\/g, '/');
const EDITOR_MODULE_URL = EDITOR_MODULE_PATH.startsWith('/')
  ? `/@fs${EDITOR_MODULE_PATH}`
  : `/@fs/${EDITOR_MODULE_PATH}`;
const VITE_URL = process.env.VITE_URL || 'http://localhost:7700';
const SAMPLE_FILE = 'biz_plan.hwp';

await runTest('field-agent exact table cell HWP apply/reopen/revert gate', async ({ page }) => {
  await page.goto(`${VITE_URL}/e2e/embed-harness.html`, { waitUntil: 'domcontentloaded' });

  const result = await page.evaluate(async ({ editorModuleUrl, sampleFile }) => {
    const { createEditor } = await import(editorModuleUrl);
    const { collectFieldTargetEvidence } = await import('/src/document-agent/controller.ts');
    const host = document.createElement('div');
    host.style.cssText = 'width: 100vw; height: 100vh';
    document.body.replaceChildren(host);
    const editor = await createEditor(host, {
      studioUrl: `${location.origin}/`,
      renderer: 'canvas2d',
      handshakeTimeoutMs: 10_000,
    });
    const sampleUrl = `/samples/${sampleFile.split('/').map(encodeURIComponent).join('/')}`;
    const bytes = await fetch(sampleUrl).then(response => response.arrayBuffer());
    await editor.loadFile(bytes, sampleFile, { suppressDialogs: true });
    const studioWindow = editor.element.contentWindow;
    const wasm = studioWindow.__wasm;
    if (!wasm) throw new Error('Studio WasmBridge is unavailable');

    const sameTarget = (left, right) => Boolean(left && right)
      && left.kind === right.kind
      && left.section === right.section
      && left.parentPara === right.parentPara
      && left.controlIndex === right.controlIndex
      && left.cellIndex === right.cellIndex
      && left.cellParagraph === right.cellParagraph;
    const readTargetText = (bridge, target) => {
      const length = bridge.getCellParagraphLength(
        target.section,
        target.parentPara,
        target.controlIndex,
        target.cellIndex,
        target.cellParagraph,
      );
      return length > 0 ? bridge.getTextInCell(
        target.section,
        target.parentPara,
        target.controlIndex,
        target.cellIndex,
        target.cellParagraph,
        0,
        length,
      ) : '';
    };
    const findCandidate = () => {
      for (let section = 0; section < wasm.getSectionCount(); section += 1) {
        for (let parentPara = 0; parentPara < wasm.getParagraphCount(section); parentPara += 1) {
          const controlCount = wasm.getControlTextPositions(section, parentPara).length;
          for (let controlIndex = 0; controlIndex < controlCount; controlIndex += 1) {
            let dimensions;
            try {
              dimensions = wasm.getTableDimensions(section, parentPara, controlIndex);
            } catch {
              continue;
            }
            for (let cellIndex = 0; cellIndex < dimensions.cellCount; cellIndex += 1) {
              const paragraphCount = wasm.getCellParagraphCount(
                section,
                parentPara,
                controlIndex,
                cellIndex,
              );
              for (let cellParagraph = 0; cellParagraph < paragraphCount; cellParagraph += 1) {
                const target = {
                  kind: 'table_cell_text',
                  section,
                  parentPara,
                  controlIndex,
                  cellIndex,
                  cellParagraph,
                };
                const text = readTargetText(wasm, target);
                const chars = Array.from(text);
                if (chars.length < 1 || chars.length > 80 || !text.trim()) continue;
                try {
                  const evidence = collectFieldTargetEvidence(wasm, target);
                  return { target, evidence };
                } catch {
                  // 혼합 서식 등 atomic field contract에 맞지 않는 셀은 건너뛴다.
                }
              }
            }
          }
        }
      }
      throw new Error(`safe table cell candidate not found: ${sampleFile}`);
    };
    const reopen = async (exportedBytes, target) => {
      const reopenHost = document.createElement('div');
      reopenHost.style.cssText = 'position: fixed; width: 1px; height: 1px; overflow: hidden';
      document.body.append(reopenHost);
      const reopened = await createEditor(reopenHost, {
        studioUrl: `${location.origin}/`,
        renderer: 'canvas2d',
        handshakeTimeoutMs: 10_000,
      });
      try {
        await reopened.loadFile(exportedBytes, 'field-agent-reopened.hwp', {
          suppressDialogs: true,
        });
        const reopenedWasm = reopened.element.contentWindow.__wasm;
        if (!reopenedWasm) throw new Error('reopened Studio WasmBridge is unavailable');
        const state = await reopened.getDocumentState();
        return {
          format: state.format,
          pageCount: state.pageCount,
          text: readTargetText(reopenedWasm, target),
          modalCount: reopened.element.contentDocument.querySelectorAll('.modal-overlay').length,
        };
      } finally {
        reopened.destroy();
        reopenHost.remove();
      }
    };

    const initialState = await editor.getDocumentState();
    if (initialState.format !== 'hwp') throw new Error(`unexpected source format: ${initialState.format}`);
    const { target, evidence } = findCandidate();
    const chars = Array.from(evidence.text);
    const changedAt = chars.findIndex(char => char.trim().length > 0);
    chars[changedAt] = chars[changedAt] === '가' ? '나' : '가';
    const replacement = chars.join('');
    const documentEvents = [];
    const selectionEvents = [];
    const offDocument = editor.onDocumentChanged(event => documentEvents.push(event));
    const offSelection = editor.onFieldSelectionChanged(event => selectionEvents.push(event));

    const focus = await editor.focusFieldTarget(target);
    const focusedSelection = await editor.getFieldSelectionContext();
    const commandId = crypto.randomUUID();
    const applyStartedAt = performance.now();
    const applied = await editor.applyFieldCommand({
      schemaVersion: 1,
      commandId,
      expectedDocumentEpoch: initialState.documentEpoch,
      expectedChangeSeq: initialState.changeSeq,
      expectedDocumentSha256: initialState.documentSha256,
      target,
      expectedBeforeSha256: evidence.textSha256,
      expectedFormatSha256: evidence.formatSha256,
      expectedAdjacentContextSha256: evidence.adjacentContextSha256,
      replacement,
    });
    const applyElapsedMs = performance.now() - applyStartedAt;
    const afterApplyState = await editor.getDocumentState();
    const afterApplyEvidence = collectFieldTargetEvidence(wasm, target);
    const afterApplySelection = await editor.getFieldSelectionContext();
    const appliedBytes = await editor.exportHwp();
    const appliedBytesSha256 = [...new Uint8Array(await crypto.subtle.digest('SHA-256', appliedBytes))]
      .map(byte => byte.toString(16).padStart(2, '0')).join('');
    const afterExportState = await editor.getDocumentState();
    const appliedReopen = await reopen(appliedBytes, target);

    if (applied.afterDocumentSha256 !== afterApplyState.documentSha256
        || applied.afterDocumentSha256 !== afterExportState.documentSha256) {
      throw new Error(`HWP receipt SHA drift: ${JSON.stringify({
        receipt: applied.afterDocumentSha256,
        afterApply: afterApplyState.documentSha256,
        exportedBytes: appliedBytesSha256,
        afterExport: afterExportState.documentSha256,
      })}`);
    }

    const reverted = await editor.revertFieldCommand({
      schemaVersion: 1,
      commandId,
      expectedDocumentEpoch: applied.documentEpoch,
      expectedChangeSeq: applied.afterChangeSeq,
      expectedAfterDocumentSha256: applied.afterDocumentSha256,
      expectedAfterSha256: applied.afterTextSha256,
    });
    const finalState = await editor.getDocumentState();
    const finalEvidence = collectFieldTargetEvidence(wasm, target);
    const restoredBytes = await editor.exportHwp();
    const restoredReopen = await reopen(restoredBytes, target);
    await new Promise(resolveFrame => requestAnimationFrame(() => resolveFrame()));

    const output = {
      sourceFormat: initialState.format,
      target,
      original: evidence.text,
      replacement,
      applyElapsedMs,
      focus,
      focusedSelection,
      afterApplySelection,
      afterApplyText: afterApplyEvidence.text,
      afterApplyFormatStable: afterApplyEvidence.formatSha256 === evidence.formatSha256,
      afterApplyContextStable:
        afterApplyEvidence.adjacentContextSha256 === evidence.adjacentContextSha256,
      appliedReopen,
      restoredText: finalEvidence.text,
      restoredFormatStable: finalEvidence.formatSha256 === evidence.formatSha256,
      restoredContextStable:
        finalEvidence.adjacentContextSha256 === evidence.adjacentContextSha256,
      restoredReopen,
      pageCountBefore: initialState.pageCount,
      pageCountAfterApply: afterApplyState.pageCount,
      pageCountAfterRevert: finalState.pageCount,
      changeSeqBefore: initialState.changeSeq,
      changeSeqAfterApply: afterApplyState.changeSeq,
      changeSeqAfterRevert: finalState.changeSeq,
      applyReceiptOperation: applied.operation,
      revertReceiptOperation: reverted.operation,
      documentEventReasons: documentEvents.map(event => event.reason),
      selectionEventMatched: selectionEvents.some(event => sameTarget(event.target, target)),
      modalCount: editor.element.contentDocument.querySelectorAll('.modal-overlay').length,
    };
    offDocument();
    offSelection();
    editor.destroy();
    return output;
  }, { editorModuleUrl: EDITOR_MODULE_URL, sampleFile: SAMPLE_FILE });

  assert(result.sourceFormat === 'hwp', '실제 .hwp source format 유지');
  assert(result.applyElapsedMs <= 3000, 'field apply와 strict render 3초 이내');
  assert(result.focus.focused, 'exact table cell focus 성공');
  assert(result.focusedSelection.editable, 'focus 직후 exact cell을 editable field로 인식');
  assert(result.focusedSelection.target?.cellIndex === result.target.cellIndex,
    'focus 직후 field selection target 동기화');
  assert(result.afterApplySelection.target?.cellIndex === result.target.cellIndex,
    'apply 직후 field selection target 유지');
  assert(result.afterApplyText === result.replacement, 'field replacement 즉시 반영');
  assert(result.afterApplyFormatStable && result.afterApplyContextStable,
    'field target 서식과 표의 비대상 문맥 보존');
  assert(result.pageCountAfterApply === result.pageCountBefore, 'field apply page count 보존');
  assert(result.appliedReopen.format === 'hwp', '적용본 HWP export/reopen format 유지');
  assert(result.appliedReopen.text === result.replacement, '적용본 HWP 재개방 값 유지');
  assert(result.appliedReopen.pageCount === result.pageCountBefore,
    '적용본 HWP 재개방 page count 유지');
  assert(result.appliedReopen.modalCount === 0, '적용본 HWP 재개방 경고 modal 0회');
  assert(result.restoredText === result.original, 'field revert로 원문 복원');
  assert(result.restoredFormatStable && result.restoredContextStable,
    'field revert 뒤 서식과 표의 비대상 문맥 복원');
  assert(result.pageCountAfterRevert === result.pageCountBefore, 'field revert page count 복원');
  assert(result.restoredReopen.format === 'hwp', '복원본 HWP export/reopen format 유지');
  assert(result.restoredReopen.text === result.original, '복원본 HWP 재개방 원문 유지');
  assert(result.changeSeqAfterApply === result.changeSeqBefore + 1,
    'field apply changeSeq 정확히 1 증가');
  assert(result.changeSeqAfterRevert === result.changeSeqBefore + 2,
    'field revert changeSeq 정확히 2 증가');
  assert(result.applyReceiptOperation === 'apply' && result.revertReceiptOperation === 'revert',
    'field apply/revert receipt terminal 상태');
  assert(JSON.stringify(result.documentEventReasons) === JSON.stringify([
    'field_agent_apply', 'field_agent_revert',
  ]), 'field public document event 순서');
  assert(result.selectionEventMatched, 'field selection event에 exact target 전달');
  assert(result.modalCount === 0, 'field focus/apply/revert 중 경고 modal 0회');
}, { skipLoadApp: true });

await runTest('field-agent exact form_text HWP apply/reopen/revert gate', async ({ page }) => {
  await page.goto(`${VITE_URL}/e2e/embed-harness.html`, { waitUntil: 'domcontentloaded' });

  const result = await page.evaluate(async ({ editorModuleUrl }) => {
    const { createEditor } = await import(editorModuleUrl);
    const { collectFieldTargetEvidence } = await import('/src/document-agent/controller.ts');
    const host = document.createElement('div');
    host.style.cssText = 'width: 100vw; height: 100vh';
    document.body.replaceChildren(host);
    const editor = await createEditor(host, {
      studioUrl: `${location.origin}/`,
      renderer: 'canvas2d',
      handshakeTimeoutMs: 10_000,
    });
    const bytes = await fetch('/samples/field-01.hwp').then(response => response.arrayBuffer());
    await editor.loadFile(bytes, 'field-01.hwp', { suppressDialogs: true });
    const wasm = editor.element.contentWindow.__wasm;
    if (!wasm) throw new Error('Studio WasmBridge is unavailable');
    const field = wasm.getFieldList().find(entry =>
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
    const initialState = await editor.getDocumentState();
    const before = collectFieldTargetEvidence(wasm, target);
    const otherField = wasm.getFieldList().find(entry =>
      entry.fieldId !== field.fieldId
      && entry.fieldType === 'clickhere'
      && entry.editableInForm === true
      && !(entry.location.path?.length));
    if (!otherField) throw new Error('second root form_text field not found');
    await editor.focusFieldTarget({
      kind: 'form_text',
      section: otherField.location.sectionIndex,
      paragraph: otherField.location.paraIndex,
      fieldId: otherField.fieldId,
    });
    await new Promise(resolveFrame => requestAnimationFrame(() => resolveFrame()));
    const documentEvents = [];
    const selectionEvents = [];
    const offDocument = editor.onDocumentChanged(event => documentEvents.push(event.reason));
    const offSelection = editor.onFieldSelectionChanged(event => selectionEvents.push(event.target));
    const focus = await editor.focusFieldTarget(target);
    await new Promise(resolveFrame => requestAnimationFrame(() => resolveFrame()));
    const focusedSelection = await editor.getFieldSelectionContext();
    const commandId = crypto.randomUUID();
    const startedAt = performance.now();
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
    const applyElapsedMs = performance.now() - startedAt;
    const after = collectFieldTargetEvidence(wasm, target);
    const afterState = await editor.getDocumentState();
    const appliedBytes = await editor.exportHwp();

    const reopenHost = document.createElement('div');
    reopenHost.style.cssText = 'position: fixed; width: 1px; height: 1px; overflow: hidden';
    document.body.append(reopenHost);
    const reopened = await createEditor(reopenHost, {
      studioUrl: `${location.origin}/`,
      renderer: 'canvas2d',
      handshakeTimeoutMs: 10_000,
    });
    await reopened.loadFile(appliedBytes, 'field-01-applied.hwp', { suppressDialogs: true });
    const reopenedWasm = reopened.element.contentWindow.__wasm;
    if (!reopenedWasm) throw new Error('reopened Studio WasmBridge is unavailable');
    const reopenedValue = reopenedWasm.getFieldValue(field.fieldId).value;
    const reopenedPageCount = (await reopened.getDocumentState()).pageCount;
    reopened.destroy();
    reopenHost.remove();

    const reverted = await editor.revertFieldCommand({
      schemaVersion: 1,
      commandId,
      expectedDocumentEpoch: applied.documentEpoch,
      expectedChangeSeq: applied.afterChangeSeq,
      expectedAfterDocumentSha256: applied.afterDocumentSha256,
      expectedAfterSha256: applied.afterTextSha256,
    });
    const restored = collectFieldTargetEvidence(wasm, target);
    const finalState = await editor.getDocumentState();
    await new Promise(resolveFrame => requestAnimationFrame(resolveFrame));
    const output = {
      target,
      original: before.text,
      focus,
      focusedSelection,
      applyElapsedMs,
      appliedText: after.text,
      formatStable: after.formatSha256 === before.formatSha256,
      contextStable: after.adjacentContextSha256 === before.adjacentContextSha256,
      reopenedValue,
      reopenedPageCount,
      restoredText: restored.text,
      restoredFormatStable: restored.formatSha256 === before.formatSha256,
      restoredContextStable: restored.adjacentContextSha256 === before.adjacentContextSha256,
      pageCounts: [initialState.pageCount, afterState.pageCount, finalState.pageCount],
      changeSeqs: [initialState.changeSeq, applied.afterChangeSeq, reverted.afterChangeSeq],
      documentEvents,
      selectionEventMatched: selectionEvents.some(selection =>
        selection?.kind === 'form_text' && selection.fieldId === target.fieldId),
      modalCount: editor.element.contentDocument.querySelectorAll('.modal-overlay').length,
    };
    offDocument();
    offSelection();
    editor.destroy();
    return output;
  }, { editorModuleUrl: EDITOR_MODULE_URL });

  assert(result.original === '', '빈 누름틀 preimage를 exact 값으로 읽음');
  assert(result.applyElapsedMs <= 3000, 'form_text apply와 strict render 3초 이내');
  assert(result.focus.focused, 'exact form_text focus 성공');
  assert(result.focusedSelection.editable, 'focus 직후 form_text를 editable field로 인식');
  assert(result.focusedSelection.target?.fieldId === result.target.fieldId,
    'focus 직후 form_text selection target 동기화');
  assert(result.appliedText === '주식회사 노튼', 'form_text replacement 즉시 반영');
  assert(result.formatStable && result.contextStable, '누름틀 서식과 비대상 문맥 보존');
  assert(result.reopenedValue === '주식회사 노튼', '적용본 HWP 재개방 누름틀 값 유지');
  assert(result.reopenedPageCount === result.pageCounts[0], '적용본 HWP 재개방 page count 유지');
  assert(result.restoredText === result.original, 'form_text revert로 원문 복원');
  assert(result.restoredFormatStable && result.restoredContextStable,
    'form_text revert 뒤 서식과 비대상 문맥 복원');
  assert(result.pageCounts.every(count => count === result.pageCounts[0]),
    'form_text apply/revert page count 보존');
  assert(result.changeSeqs[1] === result.changeSeqs[0] + 1
    && result.changeSeqs[2] === result.changeSeqs[0] + 2,
  'form_text apply/revert changeSeq 정확히 1씩 증가');
  assert(JSON.stringify(result.documentEvents) === JSON.stringify([
    'field_agent_apply', 'field_agent_revert',
  ]), 'form_text public document event 순서');
  assert(result.selectionEventMatched, 'field selection event에 exact form_text target 전달');
  assert(result.modalCount === 0, 'form_text focus/apply/revert 중 경고 modal 0회');
}, { skipLoadApp: true });
