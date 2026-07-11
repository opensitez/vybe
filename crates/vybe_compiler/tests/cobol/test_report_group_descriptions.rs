use super::helpers::compile_ok;

#[test]
fn report_heading_group_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RGD1.\nDATA DIVISION.\nREPORT SECTION.\nRD SALES-RPT.\n01 HEAD-LINE TYPE REPORT HEADING.\n   03 COL 1 VALUE \"SALES\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_detail_group_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RGD2.\nDATA DIVISION.\nREPORT SECTION.\nRD SALES-RPT.\n01 DETAIL-LINE TYPE DETAIL.\n   03 COL 1 PIC X(10) SOURCE WS-NAME.\nWORKING-STORAGE SECTION.\n01 WS-NAME PIC X(10).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_footing_group_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RGD3.\nDATA DIVISION.\nREPORT SECTION.\nRD SALES-RPT.\n01 FOOT-LINE TYPE REPORT FOOTING.\n   03 COL 1 VALUE \"END\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_column_number_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RGD4.\nDATA DIVISION.\nREPORT SECTION.\nRD SALES-RPT.\n01 DETAIL-LINE TYPE DETAIL.\n   03 COLUMN NUMBER 5 PIC X(5) VALUE \"ITEM\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_page_heading_group_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RGD5.\nDATA DIVISION.\nREPORT SECTION.\nRD SALES-RPT.\n01 PAGE-HEAD TYPE PAGE HEADING.\n   03 COL 1 VALUE \"HEADER\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_page_footing_group_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RGD6.\nDATA DIVISION.\nREPORT SECTION.\nRD SALES-RPT.\n01 PAGE-FOOT TYPE PAGE FOOTING.\n   03 COL 1 VALUE \"FOOT\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_control_heading_group_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RGD7.\nDATA DIVISION.\nREPORT SECTION.\nRD SALES-RPT CONTROLS ARE WS-DEPT.\n01 CH1 TYPE CONTROL HEADING WS-DEPT.\n   03 COL 1 VALUE \"CH\".\nWORKING-STORAGE SECTION.\n01 WS-DEPT PIC X(4).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_control_footing_group_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RGD8.\nDATA DIVISION.\nREPORT SECTION.\nRD SALES-RPT CONTROLS ARE WS-DEPT.\n01 CF1 TYPE CONTROL FOOTING WS-DEPT.\n   03 COL 1 VALUE \"CF\".\nWORKING-STORAGE SECTION.\n01 WS-DEPT PIC X(4).\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_report_footing_group_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RGD9.\nDATA DIVISION.\nREPORT SECTION.\nRD SALES-RPT.\n01 RF1 TYPE REPORT FOOTING.\n   03 COL 1 VALUE \"DONE\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_line_plus_clause_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RGD10.\nDATA DIVISION.\nREPORT SECTION.\nRD SALES-RPT.\n01 D1 TYPE DETAIL.\n   03 LINE PLUS 1 COL 1 VALUE \"ROW\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}
