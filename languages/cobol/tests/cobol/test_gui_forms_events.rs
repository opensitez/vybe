use super::helpers::compile_ok;

#[test]
fn screen_basic_display_accept_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S1.\nDATA DIVISION.\nSCREEN SECTION.\n01 SCR.\n   05 LINE 1 COLUMN 1 VALUE \"Name\".\nWORKING-STORAGE SECTION.\n01 N PIC X(20).\nPROCEDURE DIVISION.\n    DISPLAY SCR.\n    ACCEPT N.\n    STOP RUN.",
    );
}
#[test]
fn screen_menu_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S2.\nDATA DIVISION.\nSCREEN SECTION.\n01 SCR.\n   05 LINE 1 COLUMN 1 VALUE \"1. Add\".\n   05 LINE 2 COLUMN 1 VALUE \"2. Exit\".\nWORKING-STORAGE SECTION.\n01 C PIC 9.\nPROCEDURE DIVISION.\n    DISPLAY SCR.\n    ACCEPT C.\n    STOP RUN.",
    );
}
#[test]
fn screen_loop_menu_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S3.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 C PIC 9 VALUE 0.\nPROCEDURE DIVISION.\n    PERFORM UNTIL C = 2\n        ACCEPT C\n        IF C = 1 DISPLAY \"ADD\" END-IF\n    END-PERFORM.\n    STOP RUN.",
    );
}
#[test]
fn screen_with_evaluate_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S4.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 C PIC 9.\nPROCEDURE DIVISION.\n    ACCEPT C.\n    EVALUATE C\n        WHEN 1 DISPLAY \"A\"\n        WHEN 2 DISPLAY \"B\"\n        WHEN OTHER DISPLAY \"X\"\n    END-EVALUATE.\n    STOP RUN.",
    );
}
#[test]
fn form_field_move_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S5.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 NAME PIC X(20).\n01 OUT PIC X(20).\nPROCEDURE DIVISION.\n    ACCEPT NAME.\n    MOVE NAME TO OUT.\n    DISPLAY OUT.\n    STOP RUN.",
    );
}
#[test]
fn form_validation_if_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S6.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 AGE PIC 9(3).\nPROCEDURE DIVISION.\n    ACCEPT AGE.\n    IF AGE < 18 DISPLAY \"MINOR\" ELSE DISPLAY \"ADULT\" END-IF.\n    STOP RUN.",
    );
}
#[test]
fn form_call_handler_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S7.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 C PIC 9 VALUE 1.\nPROCEDURE DIVISION.\n    CALL \"FORM-HANDLER\" USING C.\n    STOP RUN.",
    );
}
#[test]
fn form_save_action_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S8.\nPROCEDURE DIVISION.\n    CALL \"FORM-SAVE\".\n    STOP RUN.",
    );
}
#[test]
fn form_load_action_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S9.\nPROCEDURE DIVISION.\n    CALL \"FORM-LOAD\".\n    STOP RUN.",
    );
}
#[test]
fn form_clear_action_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S10.\nPROCEDURE DIVISION.\n    CALL \"FORM-CLEAR\".\n    STOP RUN.",
    );
}
#[test]
fn form_error_display_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S11.\nPROCEDURE DIVISION.\n    DISPLAY \"ERROR\".\n    STOP RUN.",
    );
}
#[test]
fn form_success_display_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S12.\nPROCEDURE DIVISION.\n    DISPLAY \"SUCCESS\".\n    STOP RUN.",
    );
}
#[test]
fn screen_section_multiple_fields_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S13.\nDATA DIVISION.\nSCREEN SECTION.\n01 SCR.\n   05 LINE 1 COLUMN 1 PIC X(20) USING N1.\n   05 LINE 2 COLUMN 1 PIC X(20) USING N2.\nWORKING-STORAGE SECTION.\n01 N1 PIC X(20).\n01 N2 PIC X(20).\nPROCEDURE DIVISION.\n    DISPLAY SCR.\n    ACCEPT SCR.\n    STOP RUN.",
    );
}
#[test]
fn form_submit_with_call_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S14.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 PAY PIC X(50).\nPROCEDURE DIVISION.\n    CALL \"FORM-SUBMIT\" USING PAY.\n    STOP RUN.",
    );
}
#[test]
fn form_cancel_with_call_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S15.\nPROCEDURE DIVISION.\n    CALL \"FORM-CANCEL\".\n    STOP RUN.",
    );
}
#[test]
fn form_navigation_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S16.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 PAGE PIC 9 VALUE 1.\nPROCEDURE DIVISION.\n    IF PAGE = 1 DISPLAY \"P1\" END-IF.\n    STOP RUN.",
    );
}
#[test]
fn form_event_dispatch_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S17.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 EV PIC X(10) VALUE \"CLICK\".\nPROCEDURE DIVISION.\n    CALL \"UI-DISPATCH\" USING EV.\n    STOP RUN.",
    );
}
#[test]
fn form_repaint_call_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S18.\nPROCEDURE DIVISION.\n    CALL \"UI-REPAINT\".\n    STOP RUN.",
    );
}
#[test]
fn gui_button_click_handler_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S19.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 EV PIC X(10) VALUE \"CLICK\".\nPROCEDURE DIVISION.\n    CALL \"UI-BTN-HANDLER\" USING EV.\n    STOP RUN.",
    );
}
#[test]
fn gui_form_open_close_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S20.\nPROCEDURE DIVISION.\n    CALL \"FORM-OPEN\".\n    CALL \"FORM-CLOSE\".\n    STOP RUN.",
    );
}
#[test]
fn gui_field_validation_branch_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S21.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 NAME PIC X(20).\nPROCEDURE DIVISION.\n    ACCEPT NAME.\n    IF NAME = SPACES DISPLAY \"REQ\" ELSE DISPLAY \"OK\" END-IF.\n    STOP RUN.",
    );
}
#[test]
fn gui_navigation_next_prev_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S22.\nPROCEDURE DIVISION.\n    CALL \"UI-NEXT\".\n    CALL \"UI-PREV\".\n    STOP RUN.",
    );
}
#[test]
fn gui_modal_dialog_calls_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S23.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 MSG PIC X(30) VALUE \"HELLO\".\nPROCEDURE DIVISION.\n    CALL \"UI-MODAL\" USING MSG.\n    STOP RUN.",
    );
}
#[test]
fn gui_event_loop_dispatch_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. S24.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 I PIC 9 VALUE 0.\nPROCEDURE DIVISION.\n    PERFORM UNTIL I >= 2\n        ADD 1 TO I\n        CALL \"UI-DISPATCH\"\n    END-PERFORM.\n    STOP RUN.",
    );
}
