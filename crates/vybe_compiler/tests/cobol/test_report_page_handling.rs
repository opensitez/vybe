use super::helpers::compile_ok;

#[test]
fn report_page_limit_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RPH1.\nDATA DIVISION.\nREPORT SECTION.\nRD R1 PAGE LIMIT IS 60.\n01 PH TYPE PAGE HEADING.\n   03 COL 1 VALUE \"PAGE\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_first_and_last_detail_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RPH2.\nDATA DIVISION.\nREPORT SECTION.\nRD R1 PAGE LIMIT IS 66 FIRST DETAIL 5 LAST DETAIL 55.\n01 D1 TYPE DETAIL.\n   03 COL 1 VALUE \"X\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_page_heading_and_footing_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RPH3.\nDATA DIVISION.\nREPORT SECTION.\nRD R1.\n01 PH TYPE PAGE HEADING.\n   03 COL 1 VALUE \"H\".\n01 PF TYPE PAGE FOOTING.\n   03 COL 1 VALUE \"F\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_line_number_clause_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. RPH4.\nDATA DIVISION.\nREPORT SECTION.\nRD R1.\n01 D1 TYPE DETAIL.\n   03 LINE NUMBER 3 COL 1 VALUE \"ROW\".\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn report_next_group_clause_compiles() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. RPH5.\nDATA DIVISION.\nREPORT SECTION.\nRD R1.\n01 D1 TYPE DETAIL.\n   03 NEXT GROUP IS NEXT PAGE.\n   03 COL 1 VALUE \"ROW\".\nPROCEDURE DIVISION.\n    STOP RUN.");
}

#[test]
fn report_page_heading_clause_compiles() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. RPH6.\nDATA DIVISION.\nREPORT SECTION.\nRD R1 PAGE LIMIT IS 50.\n01 P1 TYPE PAGE HEADING.\n   03 LINE 1 COL 1 VALUE \"HEAD\".\nPROCEDURE DIVISION.\n    STOP RUN.");
}

#[test]
fn report_page_footing_clause_compiles() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. RPH7.\nDATA DIVISION.\nREPORT SECTION.\nRD R1 PAGE LIMIT IS 50.\n01 P1 TYPE PAGE FOOTING.\n   03 LINE 50 COL 1 VALUE \"FOOT\".\nPROCEDURE DIVISION.\n    STOP RUN.");
}

#[test]
fn report_report_heading_clause_compiles() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. RPH8.\nDATA DIVISION.\nREPORT SECTION.\nRD R1.\n01 H1 TYPE REPORT HEADING.\n   03 COL 1 VALUE \"TOP\".\nPROCEDURE DIVISION.\n    STOP RUN.");
}

#[test]
fn report_report_footing_clause_compiles() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. RPH9.\nDATA DIVISION.\nREPORT SECTION.\nRD R1.\n01 F1 TYPE REPORT FOOTING.\n   03 COL 1 VALUE \"END\".\nPROCEDURE DIVISION.\n    STOP RUN.");
}

#[test]
fn report_page_limit_and_controls_compiles() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. RPH10.\nDATA DIVISION.\nREPORT SECTION.\nRD R1 PAGE LIMIT IS 60 CONTROLS ARE WS-DEPT.\n01 D1 TYPE DETAIL.\n   03 COL 1 VALUE \"X\".\nWORKING-STORAGE SECTION.\n01 WS-DEPT PIC X(4).\nPROCEDURE DIVISION.\n    STOP RUN.");
}
