use super::helpers::{compile_ok, parse_ok};

#[test]
fn report_section_basic_rejected_or_supported() {
    let src = "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nREPORT SECTION.\nRD R1.\n01 D1 TYPE DETAIL.\n   03 LINE 1 COLUMN 1 VALUE \"X\".\nPROCEDURE DIVISION.\n    STOP RUN.";
    let _ = parse_ok(src);
}

#[test]
fn report_generate_statement_parses_or_rejects_cleanly() {
    let src = "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nREPORT SECTION.\nRD R1.\n01 D1 TYPE DETAIL.\n   03 LINE 1 COLUMN 1 VALUE \"X\".\nPROCEDURE DIVISION.\n    GENERATE D1.\n    STOP RUN.";
    let _ = parse_ok(src);
}

#[test]
fn report_writer_page_heading_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RW3.\nDATA DIVISION.\nREPORT SECTION.\nRD R1.\n01 P1 TYPE PAGE HEADING.\n   03 COL 1 VALUE \"HEAD\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_writer_page_footing_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RW4.\nDATA DIVISION.\nREPORT SECTION.\nRD R1.\n01 P1 TYPE PAGE FOOTING.\n   03 COL 1 VALUE \"FOOT\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_writer_control_heading_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RW5.\nDATA DIVISION.\nREPORT SECTION.\nRD R1 CONTROLS ARE WS-DEPT.\n01 C1 TYPE CONTROL HEADING WS-DEPT.\n   03 COL 1 VALUE \"C\".\nWORKING-STORAGE SECTION.\n01 WS-DEPT PIC X(4).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_writer_control_footing_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RW6.\nDATA DIVISION.\nREPORT SECTION.\nRD R1 CONTROLS ARE WS-DEPT.\n01 C1 TYPE CONTROL FOOTING WS-DEPT.\n   03 COL 1 VALUE \"T\".\nWORKING-STORAGE SECTION.\n01 WS-DEPT PIC X(4).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_writer_report_footing_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RW7.\nDATA DIVISION.\nREPORT SECTION.\nRD R1.\n01 F1 TYPE REPORT FOOTING.\n   03 COL 1 VALUE \"END\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_writer_sum_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RW8.\nDATA DIVISION.\nREPORT SECTION.\nRD R1.\n01 D1 TYPE DETAIL.\n   03 COL 1 PIC 9(5) SUM WS-AMT.\nWORKING-STORAGE SECTION.\n01 WS-AMT PIC 9(5).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_writer_page_limit_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RW9.\nDATA DIVISION.\nREPORT SECTION.\nRD R1 PAGE LIMIT IS 60.\n01 D1 TYPE DETAIL.\n   03 COL 1 VALUE \"X\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_writer_generate_detail_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RW10.\nDATA DIVISION.\nREPORT SECTION.\nRD R1.\n01 D1 TYPE DETAIL.\n   03 COL 1 VALUE \"X\".\nPROCEDURE DIVISION.\n    GENERATE D1.\n    STOP RUN.",
    );
}
