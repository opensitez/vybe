use super::helpers::compile_ok;

#[test]
fn accept_from_environment_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ENV1.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-PATH PIC X(64).\nPROCEDURE DIVISION.\n    ACCEPT WS-PATH FROM ENVIRONMENT \"PATH\".\n    STOP RUN.",
    );
}

#[test]
fn display_upon_environment_name_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ENV2.\nPROCEDURE DIVISION.\n    DISPLAY \"X\" UPON ENVIRONMENT-NAME.\n    STOP RUN.",
    );
}

#[test]
fn display_upon_environment_value_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ENV3.\nPROCEDURE DIVISION.\n    DISPLAY \"X\" UPON ENVIRONMENT-VALUE.\n    STOP RUN.",
    );
}

#[test]
fn accept_command_line_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ENV4.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-ARG PIC X(64).\nPROCEDURE DIVISION.\n    ACCEPT WS-ARG FROM COMMAND-LINE.\n    STOP RUN.",
    );
}

#[test]
fn accept_date_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ENV5.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-DATE PIC 9(8).\nPROCEDURE DIVISION.\n    ACCEPT WS-DATE FROM DATE YYYYMMDD.\n    STOP RUN.",
    );
}

#[test]
fn accept_day_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ENV6.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-DAY PIC 9(7).\nPROCEDURE DIVISION.\n    ACCEPT WS-DAY FROM DAY YYYYDDD.\n    STOP RUN.",
    );
}

#[test]
fn accept_time_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ENV7.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-TIME PIC 9(8).\nPROCEDURE DIVISION.\n    ACCEPT WS-TIME FROM TIME.\n    STOP RUN.",
    );
}

#[test]
fn display_upon_console_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ENV8.\nPROCEDURE DIVISION.\n    DISPLAY \"X\" UPON CONSOLE.\n    STOP RUN.",
    );
}

#[test]
fn accept_environment_name_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ENV9.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-NAME PIC X(64).\nPROCEDURE DIVISION.\n    ACCEPT WS-NAME FROM ENVIRONMENT-NAME.\n    STOP RUN.",
    );
}

#[test]
fn accept_environment_value_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. ENV10.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-VALUE PIC X(64).\nPROCEDURE DIVISION.\n    ACCEPT WS-VALUE FROM ENVIRONMENT-VALUE.\n    STOP RUN.",
    );
}
