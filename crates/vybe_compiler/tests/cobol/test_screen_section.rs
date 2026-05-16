use super::helpers::compile_ok;

// ── SCREEN SECTION declaration ────────────────────────────────

#[test] fn screen_section_blank() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       SCREEN SECTION.
       01 main-screen.
           05 BLANK SCREEN.
       PROCEDURE DIVISION.
           DISPLAY main-screen
           STOP RUN.
"#);
}

#[test] fn screen_section_literal_field() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       SCREEN SECTION.
       01 header-screen.
           05 BLANK SCREEN.
           05 LINE 1 COLUMN 1 VALUE "Welcome to COBOL".
       PROCEDURE DIVISION.
           DISPLAY header-screen
           STOP RUN.
"#);
}

#[test] fn screen_section_input_field() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-name PIC X(20).
       SCREEN SECTION.
       01 name-screen.
           05 BLANK SCREEN.
           05 LINE 2 COLUMN 5 VALUE "Name: ".
           05 LINE 2 COLUMN 12 PIC X(20) USING ws-name.
       PROCEDURE DIVISION.
           MOVE "Alice" TO ws-name
           DISPLAY name-screen
           STOP RUN.
"#);
}

#[test] fn screen_section_multiple_fields() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-first  PIC X(15).
       01 ws-last   PIC X(15).
       01 ws-age    PIC 99.
       SCREEN SECTION.
       01 entry-screen.
           05 BLANK SCREEN.
           05 LINE 1 COLUMN 1 VALUE "First Name: ".
           05 LINE 1 COLUMN 13 PIC X(15) USING ws-first.
           05 LINE 2 COLUMN 1 VALUE "Last Name:  ".
           05 LINE 2 COLUMN 13 PIC X(15) USING ws-last.
           05 LINE 3 COLUMN 1 VALUE "Age: ".
           05 LINE 3 COLUMN 6  PIC 99    USING ws-age.
       PROCEDURE DIVISION.
           MOVE "John" TO ws-first
           MOVE "Doe"  TO ws-last
           MOVE 30 TO ws-age
           DISPLAY entry-screen
           STOP RUN.
"#);
}

#[test] fn screen_section_highlight() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-val PIC X(10).
       SCREEN SECTION.
       01 hi-screen.
           05 LINE 1 COLUMN 1 VALUE "Field: " HIGHLIGHT.
           05 LINE 1 COLUMN 8 PIC X(10) USING ws-val HIGHLIGHT.
       PROCEDURE DIVISION.
           MOVE "test" TO ws-val
           DISPLAY hi-screen
           STOP RUN.
"#);
}

#[test] fn screen_section_reverse_video() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-title PIC X(20) VALUE "STATUS: OK".
       SCREEN SECTION.
       01 status-bar.
           05 LINE 24 COLUMN 1 PIC X(20) FROM ws-title REVERSE-VIDEO.
       PROCEDURE DIVISION.
           DISPLAY status-bar
           STOP RUN.
"#);
}

#[test] fn screen_section_blink() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       SCREEN SECTION.
       01 alert-screen.
           05 LINE 12 COLUMN 30 VALUE "ALERT!" BLINK.
       PROCEDURE DIVISION.
           DISPLAY alert-screen
           STOP RUN.
"#);
}

#[test] fn screen_section_foreground_background() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       SCREEN SECTION.
       01 colored-screen.
           05 BLANK SCREEN.
           05 LINE 1 COLUMN 1
              VALUE "Error!"
              FOREGROUND-COLOR 4
              BACKGROUND-COLOR 0.
       PROCEDURE DIVISION.
           DISPLAY colored-screen
           STOP RUN.
"#);
}

#[test] fn screen_section_auto_tab() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-code PIC X(5).
       SCREEN SECTION.
       01 code-screen.
           05 LINE 5 COLUMN 10 VALUE "Code: ".
           05 LINE 5 COLUMN 16 PIC X(5) USING ws-code AUTO.
       PROCEDURE DIVISION.
           DISPLAY code-screen
           STOP RUN.
"#);
}

#[test] fn screen_section_required() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-id PIC 9(8).
       SCREEN SECTION.
       01 id-screen.
           05 LINE 3 COLUMN 5 VALUE "ID: ".
           05 LINE 3 COLUMN 9 PIC 9(8) USING ws-id REQUIRED.
       PROCEDURE DIVISION.
           MOVE 12345678 TO ws-id
           DISPLAY id-screen
           STOP RUN.
"#);
}

#[test] fn screen_section_protected() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-readonly PIC X(20) VALUE "READ ONLY".
       SCREEN SECTION.
       01 view-screen.
           05 LINE 1 COLUMN 1 PIC X(20) FROM ws-readonly PROTECTED.
       PROCEDURE DIVISION.
           DISPLAY view-screen
           STOP RUN.
"#);
}

#[test] fn screen_section_secure_password() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-pass PIC X(16).
       SCREEN SECTION.
       01 login-screen.
           05 BLANK SCREEN.
           05 LINE 5 COLUMN 20 VALUE "Password: ".
           05 LINE 5 COLUMN 30 PIC X(16) USING ws-pass SECURE.
       PROCEDURE DIVISION.
           DISPLAY login-screen
           STOP RUN.
"#);
}

#[test] fn screen_section_with_accept() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-choice PIC X VALUE SPACE.
       SCREEN SECTION.
       01 menu-screen.
           05 BLANK SCREEN.
           05 LINE 1 COLUMN 1 VALUE "1. Option A".
           05 LINE 2 COLUMN 1 VALUE "2. Option B".
           05 LINE 4 COLUMN 1 VALUE "Choice: ".
           05 choice-fld LINE 4 COLUMN 9 PIC X USING ws-choice.
       PROCEDURE DIVISION.
           DISPLAY menu-screen
           STOP RUN.
"#);
}

#[test] fn screen_section_nested_group() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-city  PIC X(20).
       01 ws-state PIC XX.
       SCREEN SECTION.
       01 addr-screen.
           05 top-line.
               10 BLANK SCREEN.
               10 LINE 1 COLUMN 1 VALUE "Address Entry".
           05 city-grp.
               10 LINE 3 COLUMN 1 VALUE "City:  ".
               10 LINE 3 COLUMN 8 PIC X(20) USING ws-city.
               10 LINE 3 COLUMN 29 VALUE "State: ".
               10 LINE 3 COLUMN 36 PIC XX USING ws-state.
       PROCEDURE DIVISION.
           MOVE "Springfield" TO ws-city
           MOVE "IL" TO ws-state
           DISPLAY addr-screen
           STOP RUN.
"#);
}

#[test] fn screen_section_grid_layout() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-vals.
           05 ws-val-1 PIC 999 VALUE 100.
           05 ws-val-2 PIC 999 VALUE 200.
           05 ws-val-3 PIC 999 VALUE 300.
       SCREEN SECTION.
       01 grid-screen.
           05 BLANK SCREEN.
           05 LINE 1 COLUMN 1  VALUE "Col1".
           05 LINE 1 COLUMN 10 VALUE "Col2".
           05 LINE 1 COLUMN 20 VALUE "Col3".
           05 LINE 2 COLUMN 1  PIC 999 FROM ws-val-1.
           05 LINE 2 COLUMN 10 PIC 999 FROM ws-val-2.
           05 LINE 2 COLUMN 20 PIC 999 FROM ws-val-3.
       PROCEDURE DIVISION.
           DISPLAY grid-screen
           STOP RUN.
"#);
}
