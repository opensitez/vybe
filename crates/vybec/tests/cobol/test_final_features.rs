use vybec::parser_cobol::parse;
use vybec::compiler_cobol::Compiler;

fn compile_ok(src: &str) {
    let program = parse(src).expect("parse failed");
    let c = Compiler::new();
    let res = c.compile(&program);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

fn p(data: &str, body: &str) -> String {
    format!("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.", data, body)
}

// ═══════════════════════════════════════════════════════════
// 1. COPY REPLACING
// ═══════════════════════════════════════════════════════════
#[test] fn copy_basic() { compile_ok(&p("", "    COPY COMMON-DEFS.")); }
#[test] fn copy_replacing() { compile_ok(&p("", "    COPY CUSTOMER-REC REPLACING OLD-NAME BY NEW-NAME.")); }
#[test] fn copy_replacing_multi() { compile_ok(&p("", "    COPY RECORD-DEF REPLACING \"OLD\" BY \"NEW\" \"FIELD1\" BY \"FIELD2\".")); }

// ═══════════════════════════════════════════════════════════
// 2. FILE SECTION with FD/SD
// ═══════════════════════════════════════════════════════════
#[test]
fn fd_basic() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. FDTEST.
DATA DIVISION.
FILE SECTION.
FD CUSTOMER-FILE RECORD CONTAINS 80 CHARACTERS.
01 CUSTOMER-RECORD.
   05 CUST-ID   PIC 9(5).
   05 CUST-NAME PIC X(30).
   05 CUST-BAL  PIC 9(8)V99.
WORKING-STORAGE SECTION.
01 WS-REC PIC X(80).
PROCEDURE DIVISION.
    DISPLAY "FD Test".
    STOP RUN.
"#);
}

#[test]
fn sd_sort_file() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. SDTEST.
DATA DIVISION.
FILE SECTION.
SD SORT-FILE.
01 SORT-RECORD.
   05 SORT-KEY PIC 9(5).
   05 SORT-DATA PIC X(75).
WORKING-STORAGE SECTION.
01 WS-REC PIC X(80).
PROCEDURE DIVISION.
    DISPLAY "SD Test".
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// 3. PERFORM WITH TEST AFTER (do-while)
// ═══════════════════════════════════════════════════════════
#[test]
fn perform_test_after() {
    compile_ok(&p(
        "01 WS-I PIC 9(3) VALUE 0.",
        "    PERFORM WITH TEST AFTER UNTIL WS-I >= 5\n        ADD 1 TO WS-I\n        DISPLAY WS-I\n    END-PERFORM."
    ));
}

#[test]
fn perform_test_before() {
    compile_ok(&p(
        "01 WS-I PIC 9(3) VALUE 0.",
        "    PERFORM WITH TEST BEFORE UNTIL WS-I >= 5\n        ADD 1 TO WS-I\n    END-PERFORM."
    ));
}

#[test]
fn perform_test_after_runs_once() {
    compile_ok(&p(
        "01 WS-I PIC 9(3) VALUE 10.",
        "    PERFORM WITH TEST AFTER UNTIL WS-I >= 5\n        DISPLAY \"Ran at least once\"\n        ADD 1 TO WS-I\n    END-PERFORM."
    ));
}

// ═══════════════════════════════════════════════════════════
// 4. STRING WITH POINTER
// ═══════════════════════════════════════════════════════════
#[test]
fn string_with_pointer() {
    compile_ok(&p(
        "01 WS-A PIC X(10) VALUE \"Hello\".\n01 WS-B PIC X(10) VALUE \"World\".\n01 WS-R PIC X(25).\n01 WS-PTR PIC 9(3) VALUE 1.",
        "    STRING WS-A DELIMITED BY SIZE WS-B DELIMITED BY SIZE INTO WS-R WITH POINTER WS-PTR."
    ));
}

// ═══════════════════════════════════════════════════════════
// 5. UNSTRING with COUNT/POINTER
// ═══════════════════════════════════════════════════════════
#[test]
fn unstring_with_count() {
    compile_ok(&p(
        "01 WS-SRC PIC X(30) VALUE \"A,BB,CCC\".\n01 F1 PIC X(10).\n01 F2 PIC X(10).\n01 F3 PIC X(10).\n01 C1 PIC 9(3).\n01 C2 PIC 9(3).\n01 C3 PIC 9(3).",
        "    UNSTRING WS-SRC DELIMITED BY \",\" INTO F1 COUNT C1 F2 COUNT C2 F3 COUNT C3."
    ));
}

#[test]
fn unstring_multi_delim() {
    compile_ok(&p(
        "01 WS-SRC PIC X(30) VALUE \"A,B;C\".\n01 F1 PIC X(10).\n01 F2 PIC X(10).\n01 F3 PIC X(10).",
        "    UNSTRING WS-SRC DELIMITED BY \",\" OR \";\" INTO F1 F2 F3."
    ));
}

// ═══════════════════════════════════════════════════════════
// 6. OCCURS DEPENDING ON
// ═══════════════════════════════════════════════════════════
#[test]
fn occurs_depending_on() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. ODO.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-COUNT PIC 9(3) VALUE 5.
01 WS-TABLE.
   05 WS-ITEM PIC X(10) OCCURS 1 TO 100 TIMES
      DEPENDING ON WS-COUNT.
PROCEDURE DIVISION.
    DISPLAY "ODO Test".
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// 7. 88-level VALUE THRU
// ═══════════════════════════════════════════════════════════
#[test]
fn cond_88_range() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. RANGE88.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-AGE PIC 9(3) VALUE 25.
   88 IS-CHILD  VALUE 0 THRU 12.
   88 IS-TEEN   VALUE 13 THRU 19.
   88 IS-ADULT  VALUE 20 THRU 120.
PROCEDURE DIVISION.
    IF IS-ADULT
        DISPLAY "Adult"
    END-IF.
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// 8. ADD/SUBTRACT CORRESPONDING
// ═══════════════════════════════════════════════════════════
#[test]
fn add_corresponding() {
    compile_ok(&p(
        "01 SRC.\n   05 AMT PIC 9(5) VALUE 100.\n01 DST.\n   05 AMT PIC 9(5) VALUE 50.",
        "    ADD CORRESPONDING SRC TO DST."
    ));
}

#[test]
fn subtract_corresponding() {
    compile_ok(&p(
        "01 SRC.\n   05 AMT PIC 9(5) VALUE 30.\n01 DST.\n   05 AMT PIC 9(5) VALUE 100.",
        "    SUBTRACT CORRESPONDING SRC FROM DST."
    ));
}

#[test]
fn add_corr() {
    compile_ok(&p(
        "01 A.\n   05 X PIC 9(5) VALUE 10.\n01 B.\n   05 X PIC 9(5) VALUE 20.",
        "    ADD CORR A TO B."
    ));
}

// ═══════════════════════════════════════════════════════════
// 9. ACCEPT FROM COMMAND-LINE
// ═══════════════════════════════════════════════════════════
#[test]
fn accept_command_line() {
    compile_ok(&p(
        "01 WS-ARGS PIC X(100).",
        "    ACCEPT WS-ARGS FROM COMMAND-LINE."
    ));
}

// ═══════════════════════════════════════════════════════════
// 10. SET 88-level TO TRUE
// ═══════════════════════════════════════════════════════════
#[test]
fn set_88_true() {
    compile_ok(&p(
        "01 WS-STATUS PIC X(1).\n   88 IS-ACTIVE VALUE \"A\".\n   88 IS-INACTIVE VALUE \"I\".",
        "    SET IS-ACTIVE TO TRUE.\n    SET IS-INACTIVE TO FALSE."
    ));
}

// ═══════════════════════════════════════════════════════════
// 11. SPECIAL-NAMES
// ═══════════════════════════════════════════════════════════
#[test]
fn special_names_decimal_comma() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. SPECNAMES.
ENVIRONMENT DIVISION.
SPECIAL-NAMES.
    DECIMAL-POINT IS COMMA.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-AMT PIC 9(5)V99 VALUE 1234.56.
PROCEDURE DIVISION.
    DISPLAY WS-AMT.
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// 12. EVALUATE ALSO
// ═══════════════════════════════════════════════════════════
#[test]
fn evaluate_nested_for_also() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. EVALALSO.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STATUS PIC X(1) VALUE "A".
01 WS-REGION PIC 9(1) VALUE 1.
PROCEDURE DIVISION.
    EVALUATE WS-STATUS
        WHEN "A"
            EVALUATE WS-REGION
                WHEN 1
                    DISPLAY "Active Region 1"
                WHEN 2
                    DISPLAY "Active Region 2"
            END-EVALUATE
        WHEN "I"
            DISPLAY "Inactive"
    END-EVALUATE.
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// 13. UTF-8/NATIONAL type
// ═══════════════════════════════════════════════════════════
#[test]
fn national_type() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. UTF8.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(30) NATIONAL VALUE "Unicode Test".
PROCEDURE DIVISION.
    DISPLAY WS-NAME.
    STOP RUN.
"#);
}

#[test]
fn national_usage() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. NATUSAGE.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(50) USAGE NATIONAL.
PROCEDURE DIVISION.
    MOVE "Hello World" TO WS-TEXT.
    DISPLAY WS-TEXT.
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// COMPLEX PROGRAMS WITH NEW FEATURES
// ═══════════════════════════════════════════════════════════
#[test]
fn do_while_menu() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. MENU.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-CHOICE PIC 9(1) VALUE 0.
PROCEDURE DIVISION.
    PERFORM WITH TEST AFTER UNTIL WS-CHOICE = 9
        DISPLAY "1. Option A"
        DISPLAY "2. Option B"
        DISPLAY "9. Exit"
        MOVE 9 TO WS-CHOICE
    END-PERFORM.
    DISPLAY "Goodbye".
    STOP RUN.
"#);
}

#[test]
fn file_section_with_processing() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. FILEPROC.
DATA DIVISION.
FILE SECTION.
FD INPUT-FILE RECORD CONTAINS 100 CHARACTERS.
01 INPUT-RECORD.
   05 REC-ID    PIC 9(5).
   05 REC-NAME  PIC X(30).
   05 REC-AMT   PIC 9(8)V99.
FD OUTPUT-FILE RECORD CONTAINS 100 CHARACTERS.
01 OUTPUT-RECORD.
   05 OUT-ID    PIC 9(5).
   05 OUT-NAME  PIC X(30).
   05 OUT-AMT   PIC 9(8)V99.
WORKING-STORAGE SECTION.
01 WS-EOF PIC 9(1) VALUE 0.
PROCEDURE DIVISION.
    DISPLAY "File processing".
    STOP RUN.
"#);
}

#[test]
fn corresponding_arithmetic() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. CORRARITH.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-CURRENT.
   05 WS-SALES PIC 9(8)V99 VALUE 1000.00.
   05 WS-COSTS PIC 9(8)V99 VALUE 500.00.
01 WS-TOTAL.
   05 WS-SALES PIC 9(10)V99 VALUE 0.
   05 WS-COSTS PIC 9(10)V99 VALUE 0.
PROCEDURE DIVISION.
    ADD CORRESPONDING WS-CURRENT TO WS-TOTAL.
    ADD CORRESPONDING WS-CURRENT TO WS-TOTAL.
    ADD CORRESPONDING WS-CURRENT TO WS-TOTAL.
    DISPLAY "Total Sales: " WS-SALES OF WS-TOTAL.
    DISPLAY "Total Costs: " WS-COSTS OF WS-TOTAL.
    STOP RUN.
"#);
}

#[test]
fn command_line_program() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. CLIAPP.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ARGS PIC X(200).
01 WS-NAME PIC X(50).
PROCEDURE DIVISION.
    ACCEPT WS-ARGS FROM COMMAND-LINE.
    DISPLAY "Arguments: " WS-ARGS.
    ACCEPT WS-NAME.
    DISPLAY "Hello " WS-NAME.
    STOP RUN.
"#);
}
