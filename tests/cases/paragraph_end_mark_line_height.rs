//! 문단 끝 표시(글 뒤 빈 run)의 글자 크기가 다시 계산한 마지막 줄 높이에 들어간다.
//!
//! 한컴 저장본은 줄의 높이 · 글 높이 · 기준선 · 줄 간격을 그 줄의 최대 글자 크기로 정하고, 마지막
//! 줄에서는 문단 끝 표시의 글자 크기도 포함한다(samples HWPX 283개 실측: 끝 run이 글보다 큰 한 줄
//! 문단 227개 모두 textheight=끝 run 크기, 작은 1204개는 글 크기). rhwp 재조판은 글자 token만
//! 보아 표 칸을 되돌린 뒤 11pt 끝 run 문단의 줄 높이가 1100→1000 으로 줄었다.
#![cfg(not(target_arch = "wasm32"))]

use rhwp::model::control::Control;
use rhwp::model::paragraph::Paragraph;
use rhwp::wasm_api::HwpDocument;

const TEXT: &str = "결과물의 상용화 계획";

struct Cell {
    para: usize,
    control: usize,
}

fn table_doc() -> (HwpDocument, Cell) {
    let mut doc = HwpDocument::create_empty();
    doc.create_blank_document_native().expect("빈 문서");
    let created = doc.create_table_native(0, 0, 0, 1, 1).expect("표");
    let json: serde_json::Value = serde_json::from_str(&created).expect("표 결과 JSON");
    let cell = Cell {
        para: json["paraIdx"].as_u64().expect("paraIdx") as usize,
        control: json["controlIdx"].as_u64().expect("controlIdx") as usize,
    };
    doc.insert_text_in_cell_native(0, cell.para, cell.control, 0, 0, 0, TEXT)
        .expect("글 넣기");
    (doc, cell)
}

fn cell_paragraph<'a>(doc: &'a HwpDocument, cell: &Cell) -> &'a Paragraph {
    match &doc.document().sections[0].paragraphs[cell.para].controls[cell.control] {
        Control::Table(table) => &table.cells[0].paragraphs[0],
        _ => panic!("표가 아님"),
    }
}

fn heights(doc: &HwpDocument, cell: &Cell) -> Vec<(i32, i32, i32)> {
    cell_paragraph(doc, cell)
        .line_segs
        .iter()
        .map(|seg| (seg.line_height, seg.text_height, seg.baseline_distance))
        .collect()
}

/// 글 뒤에 보조 글자를 넣어 크기를 바꾼 뒤 지우면 그 모양이 문단 끝 빈 run으로 남는다.
fn add_trailing_run(doc: &mut HwpDocument, cell: &Cell, font_size: u32) {
    let len = TEXT.chars().count();
    doc.insert_text_in_cell_native(0, cell.para, cell.control, 0, 0, len, " ")
        .expect("보조 글자");
    doc.apply_char_format_in_cell_native(
        0,
        cell.para,
        cell.control,
        0,
        0,
        len,
        len + 1,
        &format!("{{\"fontSize\":{font_size}}}"),
    )
    .expect("보조 글자 크기");
    doc.delete_text_in_cell_native(0, cell.para, cell.control, 0, 0, len, 1)
        .expect("보조 글자 지우기");
    let shapes = &cell_paragraph(doc, cell).char_shapes;
    assert!(
        shapes.len() >= 2
            && shapes.last().unwrap().start_pos as usize == TEXT.encode_utf16().count(),
        "글 뒤 빈 run이 남아야 한다: {shapes:?}"
    );
}

#[test]
fn larger_trailing_run_sets_the_last_line_height() {
    let (mut doc, cell) = table_doc();
    let text_only = heights(&doc, &cell);
    assert_eq!(text_only, vec![(1000, 1000, 850)], "기본 10pt 글");
    add_trailing_run(&mut doc, &cell, 1100);
    assert_eq!(heights(&doc, &cell), vec![(1100, 1100, 935)]);

    // 글을 고쳐 다시 조판해도, 저장 · 다시 열기 뒤에도 끝 run 크기가 남는다.
    doc.insert_text_in_cell_native(0, cell.para, cell.control, 0, 0, 0, "※ ")
        .expect("앞에 글 넣기");
    assert_eq!(heights(&doc, &cell), vec![(1100, 1100, 935)]);
    let reopened =
        HwpDocument::from_bytes(&doc.export_hwp().expect("HWP 저장")).expect("다시 열기");
    assert_eq!(heights(&reopened, &cell), vec![(1100, 1100, 935)]);
}

/// 글자 서식 적용은 편집 prefix 없이 문단 전체를 다시 조판한다(frame 경로).
#[test]
fn full_reflow_keeps_the_end_mark_size() {
    let (mut doc, cell) = table_doc();
    add_trailing_run(&mut doc, &cell, 1100);
    let len = TEXT.chars().count();
    doc.apply_char_format_in_cell_native(
        0,
        cell.para,
        cell.control,
        0,
        0,
        0,
        len,
        "{\"bold\":true}",
    )
    .expect("글 굵게");
    assert_eq!(heights(&doc, &cell), vec![(1100, 1100, 935)]);
}

#[test]
fn smaller_trailing_run_keeps_the_text_height() {
    let (mut doc, cell) = table_doc();
    add_trailing_run(&mut doc, &cell, 800);
    assert_eq!(heights(&doc, &cell), vec![(1000, 1000, 850)]);
}

/// 본문 둘째 문단(구역 정의 같은 컨트롤이 없어 frame 재조판 대상)에 글을 넣는다.
fn body_doc(text: &str) -> HwpDocument {
    let mut doc = HwpDocument::create_empty();
    doc.create_blank_document_native().expect("빈 문서");
    doc.insert_text_native(0, 0, 0, "지원서").expect("첫 문단");
    doc.split_paragraph_native(0, 0, 3, None)
        .expect("문단 나누기");
    doc.insert_text_native(0, 1, 0, text).expect("본문 글");
    assert!(doc.document().sections[0].paragraphs[1].controls.is_empty());
    doc
}

fn body_heights(doc: &HwpDocument) -> Vec<(i32, i32, i32)> {
    doc.document().sections[0].paragraphs[1]
        .line_segs
        .iter()
        .map(|seg| (seg.line_height, seg.text_height, seg.baseline_distance))
        .collect()
}

/// 본문 문단 끝 빈 run을 만들고 글 전체에 서식을 줘 문단 전체를 다시 조판한다.
fn body_trailing_run_then_full_reflow(doc: &mut HwpDocument, text: &str, font_size: u32) {
    let len = text.chars().count();
    doc.insert_text_native(0, 1, len, " ").expect("보조 글자");
    doc.apply_char_format_native(0, 1, len, len + 1, &format!("{{\"fontSize\":{font_size}}}"))
        .expect("보조 글자 크기");
    doc.delete_text_native(0, 1, len, 1)
        .expect("보조 글자 지우기");
    let para = &doc.document().sections[0].paragraphs[1];
    assert_eq!(
        para.char_shapes
            .last()
            .map(|shape| shape.start_pos as usize),
        Some(text.encode_utf16().count()),
        "글 뒤 빈 run(UTF-16 위치)이 남아야 한다: {:?}",
        para.char_shapes
    );
    doc.apply_char_format_native(0, 1, 0, len, "{\"bold\":true}")
        .expect("글 굵게");
}

#[test]
fn body_paragraph_with_emoji_uses_the_end_mark_size() {
    // 🚀 은 UTF-16 두 단위다 — 끝 run 위치는 UTF-16, 줄 나눔은 글자 단위.
    let text = "기대성과 🚀 상용화 계획";
    let mut doc = body_doc(text);
    body_trailing_run_then_full_reflow(&mut doc, text, 1100);
    assert_eq!(body_heights(&doc), vec![(1100, 1100, 935)]);
}

#[test]
fn only_the_last_body_line_takes_the_end_mark_size() {
    let text = "시제품을 OTT 플랫폼에 공급해 첫해 매출 3억 원을 만듭니다. ".repeat(6);
    let text = text.trim_end();
    let mut doc = body_doc(text);
    body_trailing_run_then_full_reflow(&mut doc, text, 1400);
    let lines = body_heights(&doc);
    assert!(lines.len() > 1, "긴 본문은 여러 줄이어야 한다: {lines:?}");
    assert!(
        lines[..lines.len() - 1]
            .iter()
            .all(|line| *line == (1000, 1000, 850)),
        "{lines:?}"
    );
    assert_eq!(lines.last(), Some(&(1400, 1400, 1190)));
}
