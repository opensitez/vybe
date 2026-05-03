use super::helpers::{compile_ok, parse_ok, compile_ok_check};



fn p(data: &str, body: &str) -> String {
    format!("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.", data, body)
}

fn d() -> &'static str { "01 R PIC 9(10) VALUE 0.\n01 A PIC 9(10) VALUE 10.\n01 B PIC 9(10) VALUE 20." }

// ═══════════════════════════════════════════════════════════
// EXIT PERFORM / EXIT PARAGRAPH
// ═══════════════════════════════════════════════════════════
#[test] fn exit_perform() { compile_ok(&p(d(), "    PERFORM 10 TIMES\n        EXIT PERFORM\n    END-PERFORM.")); }
#[test] fn exit_paragraph() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    PERFORM MY-PARA.\n    STOP RUN.\nMY-PARA.\n    DISPLAY \"Start\".\n    EXIT PARAGRAPH.\n    DISPLAY \"Never\".");
}
#[test] fn exit_alone() { compile_ok(&p("", "    EXIT.")); }

// ═══════════════════════════════════════════════════════════
// REWRITE / DELETE / START
// ═══════════════════════════════════════════════════════════
#[test] fn rewrite_basic() { compile_ok(&p("01 REC PIC X(80).", "    REWRITE REC.")); }
#[test] fn rewrite_from() { compile_ok(&p("01 REC PIC X(80).\n01 NEW-REC PIC X(80) VALUE \"Updated\".", "    REWRITE REC FROM NEW-REC.")); }
#[test] fn delete_file() { compile_ok(&p("", "    DELETE WS-FILE.")); }
#[test] fn start_file() { compile_ok(&p("", "    START WS-FILE KEY = WS-KEY.")); }
#[test] fn start_no_key() { compile_ok(&p("", "    START WS-FILE.")); }

// ═══════════════════════════════════════════════════════════
// INSPECT CONVERTING
// ═══════════════════════════════════════════════════════════
#[test] fn inspect_converting() { compile_ok(&p(
    "01 TXT PIC X(20) VALUE \"Hello World\".",
    "    INSPECT TXT CONVERTING \"abcdefghij\" TO \"ABCDEFGHIJ\"."
)); }
#[test] fn inspect_converting_spaces() { compile_ok(&p(
    "01 TXT PIC X(20) VALUE \"a b c\".",
    "    INSPECT TXT CONVERTING \" \" TO \"-\"."
)); }

// ═══════════════════════════════════════════════════════════
// CLASS CONDITIONS (IS NUMERIC, IS ALPHABETIC)
// ═══════════════════════════════════════════════════════════
#[test] fn is_numeric() { compile_ok(&p(
    "01 X PIC X(10) VALUE \"12345\".",
    "    IF X IS NUMERIC\n        DISPLAY \"Number\"\n    END-IF."
)); }
#[test] fn is_not_numeric() { compile_ok(&p(
    "01 X PIC X(10) VALUE \"Hello\".",
    "    IF X IS NOT NUMERIC\n        DISPLAY \"Not a number\"\n    END-IF."
)); }
#[test] fn is_alphabetic() { compile_ok(&p(
    "01 X PIC X(10) VALUE \"Hello\".",
    "    IF X IS ALPHABETIC\n        DISPLAY \"Alpha\"\n    END-IF."
)); }
#[test] fn is_alphabetic_lower() { compile_ok(&p(
    "01 X PIC X(10) VALUE \"hello\".",
    "    IF X IS ALPHABETIC-LOWER\n        DISPLAY \"Lower\"\n    END-IF."
)); }
#[test] fn is_alphabetic_upper() { compile_ok(&p(
    "01 X PIC X(10) VALUE \"HELLO\".",
    "    IF X IS ALPHABETIC-UPPER\n        DISPLAY \"Upper\"\n    END-IF."
)); }

// ═══════════════════════════════════════════════════════════
// SIGN CONDITIONS (IS POSITIVE, IS NEGATIVE, IS ZERO)
// ═══════════════════════════════════════════════════════════
#[test] fn is_positive() { compile_ok(&p(
    "01 X PIC S9(5) VALUE 10.",
    "    IF X IS POSITIVE\n        DISPLAY \"Positive\"\n    END-IF."
)); }
#[test] fn is_negative() { compile_ok(&p(
    "01 X PIC S9(5) VALUE -5.",
    "    IF X IS NEGATIVE\n        DISPLAY \"Negative\"\n    END-IF."
)); }
#[test] fn is_zero() { compile_ok(&p(
    "01 X PIC 9(5) VALUE 0.",
    "    IF X IS ZERO\n        DISPLAY \"Zero\"\n    END-IF."
)); }
#[test] fn is_not_zero() { compile_ok(&p(
    "01 X PIC 9(5) VALUE 5.",
    "    IF X IS NOT ZERO\n        DISPLAY \"Non-zero\"\n    END-IF."
)); }

// ═══════════════════════════════════════════════════════════
// COPY
// ═══════════════════════════════════════════════════════════
#[test] fn copy_stmt() { compile_ok(&p("", "    COPY COMMON-DEFS.")); }

// ═══════════════════════════════════════════════════════════
// MERGE
// ═══════════════════════════════════════════════════════════
#[test] fn merge_ascending() { compile_ok(&p("", "    MERGE WS-FILE ON ASCENDING KEY WS-KEY.")); }
#[test] fn merge_descending() { compile_ok(&p("", "    MERGE WS-FILE ON DESCENDING KEY WS-KEY.")); }

// ═══════════════════════════════════════════════════════════
// OO COBOL 2023 — INVOKE
// ═══════════════════════════════════════════════════════════
#[test] fn invoke_basic() { compile_ok(&p("01 OBJ PIC X(10).\n01 RES PIC X(10).", "    INVOKE OBJ GET-NAME RETURNING RES.")); }
#[test] fn invoke_using() { compile_ok(&p("01 OBJ PIC X(10).\n01 ARG PIC X(10) VALUE \"Test\".", "    INVOKE OBJ PROCESS USING ARG.")); }

// ═══════════════════════════════════════════════════════════
// VALIDATE / FREE / ALLOCATE
// ═══════════════════════════════════════════════════════════
#[test] fn validate_stmt() { compile_ok(&p("01 X PIC 9(5) VALUE 123.", "    VALIDATE X.")); }
#[test] fn free_stmt() { compile_ok(&p("01 PTR PIC X(10).", "    FREE PTR.")); }
#[test] fn allocate_stmt() { compile_ok(&p("01 PTR PIC X(10).", "    ALLOCATE PTR.")); }

// ═══════════════════════════════════════════════════════════
// PERFORM THRU (extended tests)
// ═══════════════════════════════════════════════════════════
#[test] fn perform_thru_3() {
    compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    PERFORM STEP-A THRU STEP-C.\n    STOP RUN.\nSTEP-A.\n    DISPLAY \"A\".\nSTEP-B.\n    DISPLAY \"B\".\nSTEP-C.\n    DISPLAY \"C\".");
}

// ═══════════════════════════════════════════════════════════
// COMPLEX PROGRAMS WITH NEW FEATURES
// ═══════════════════════════════════════════════════════════
#[test]
fn data_validation() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. VALIDATE.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-INPUT  PIC X(20) VALUE "12345".
01 WS-NAME   PIC X(20) VALUE "John Doe".
01 WS-AMOUNT PIC S9(5)V99 VALUE 150.50.
PROCEDURE DIVISION.
    IF WS-INPUT IS NUMERIC
        DISPLAY "Input is numeric"
    ELSE
        DISPLAY "Input is not numeric"
    END-IF.
    IF WS-NAME IS ALPHABETIC
        DISPLAY "Name is alpha"
    END-IF.
    IF WS-AMOUNT IS POSITIVE
        DISPLAY "Amount is positive"
    END-IF.
    IF WS-AMOUNT IS NOT ZERO
        DISPLAY "Amount is non-zero"
    END-IF.
    STOP RUN.
"#);
}

#[test]
fn char_translation() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. CONVERT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(30) VALUE "Hello World 123".
01 WS-COPY PIC X(30).
PROCEDURE DIVISION.
    MOVE WS-TEXT TO WS-COPY.
    INSPECT WS-COPY CONVERTING "abcdefghijklmnopqrstuvwxyz"
                    TO "ABCDEFGHIJKLMNOPQRSTUVWXYZ".
    DISPLAY "Original:  " WS-TEXT.
    DISPLAY "Converted: " WS-COPY.
    STOP RUN.
"#);
}

#[test]
fn exit_perform_program() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. EXITPERF.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-I PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 100
        IF WS-I = 50
            EXIT PERFORM
        END-IF
        DISPLAY WS-I
    END-PERFORM.
    DISPLAY "Finished at " WS-I.
    STOP RUN.
"#);
}

#[test]
fn file_operations() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. FILEOPS.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RECORD PIC X(80).
PROCEDURE DIVISION.
    OPEN OUTPUT WS-FILE.
    MOVE "Hello World" TO WS-RECORD.
    WRITE WS-RECORD.
    CLOSE WS-FILE.
    OPEN INPUT WS-FILE.
    READ WS-FILE INTO WS-RECORD.
    DISPLAY WS-RECORD.
    CLOSE WS-FILE.
    STOP RUN.
"#);
}

#[test]
fn invoke_oo_program() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. OOPROG.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-OBJ    PIC X(20).
01 WS-RESULT PIC X(50).
01 WS-ARG    PIC X(20) VALUE "Test Data".
PROCEDURE DIVISION.
    INVOKE WS-OBJ PROCESS USING WS-ARG RETURNING WS-RESULT.
    DISPLAY WS-RESULT.
    STOP RUN.
"#);
}

#[test]
fn multi_condition_evaluate() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. MULTIEVAL.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STATUS PIC X(1) VALUE "A".
01 WS-REGION PIC 9(1) VALUE 1.
PROCEDURE DIVISION.
    EVALUATE WS-STATUS
        WHEN "A"
            EVALUATE WS-REGION
                WHEN 1
                    DISPLAY "Active - Region 1"
                WHEN 2
                    DISPLAY "Active - Region 2"
                WHEN OTHER
                    DISPLAY "Active - Other Region"
            END-EVALUATE
        WHEN "I"
            DISPLAY "Inactive"
        WHEN OTHER
            DISPLAY "Unknown"
    END-EVALUATE.
    STOP RUN.
"#);
}

#[test]
fn batch_processing_pattern() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. BATCH.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RECORDS PIC 9(5) VALUE 0.
01 WS-ERRORS  PIC 9(5) VALUE 0.
01 WS-SUCCESS PIC 9(5) VALUE 0.
01 WS-I       PIC 9(5) VALUE 0.
01 WS-MOD     PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 100
        ADD 1 TO WS-RECORDS
        COMPUTE WS-MOD = FUNCTION MOD(WS-I 7)
        IF WS-MOD = 0
            ADD 1 TO WS-ERRORS
        ELSE
            ADD 1 TO WS-SUCCESS
        END-IF
    END-PERFORM.
    DISPLAY "Total Records: " WS-RECORDS.
    DISPLAY "Successful:    " WS-SUCCESS.
    DISPLAY "Errors:        " WS-ERRORS.
    STOP RUN.
"#);
}

#[test]
fn string_manipulation_advanced() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. STRMANIP.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME     PIC X(30) VALUE "Smith, John A.".
01 WS-LAST     PIC X(15).
01 WS-FIRST    PIC X(15).
01 WS-MIDDLE   PIC X(5).
01 WS-REVERSED PIC X(30).
01 WS-UPPER    PIC X(30).
01 WS-LEN      PIC 9(3).
01 WS-COUNT    PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    UNSTRING WS-NAME DELIMITED BY ", "
        INTO WS-LAST WS-FIRST.
    MOVE FUNCTION REVERSE(WS-NAME) TO WS-REVERSED.
    MOVE FUNCTION UPPER-CASE(WS-NAME) TO WS-UPPER.
    MOVE FUNCTION LENGTH(WS-NAME) TO WS-LEN.
    INSPECT WS-NAME TALLYING WS-COUNT FOR ALL ".".
    DISPLAY "Last:     " WS-LAST.
    DISPLAY "First:    " WS-FIRST.
    DISPLAY "Reversed: " WS-REVERSED.
    DISPLAY "Upper:    " WS-UPPER.
    DISPLAY "Length:   " WS-LEN.
    DISPLAY "Periods:  " WS-COUNT.
    STOP RUN.
"#);
}

#[test]
fn interest_calculator() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. INTEREST.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PRINCIPAL PIC 9(8)V99 VALUE 10000.00.
01 WS-RATE      PIC 9(2)V99 VALUE 5.50.
01 WS-YEARS     PIC 9(3)    VALUE 10.
01 WS-INTEREST  PIC 9(10)V99 VALUE 0.
01 WS-TOTAL     PIC 9(10)V99 VALUE 0.
01 WS-I         PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    MOVE WS-PRINCIPAL TO WS-TOTAL.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-YEARS
        COMPUTE WS-INTEREST = WS-TOTAL * (WS-RATE / 100)
        ADD WS-INTEREST TO WS-TOTAL
    END-PERFORM.
    DISPLAY "Principal: " WS-PRINCIPAL.
    DISPLAY "Rate:      " WS-RATE "%".
    DISPLAY "Years:     " WS-YEARS.
    DISPLAY "Total:     " WS-TOTAL.
    STOP RUN.
"#);
}

#[test]
fn paragraph_with_conditions() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. PARACONDN.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TYPE   PIC X(1) VALUE "A".
   88 IS-TYPE-A VALUE "A".
   88 IS-TYPE-B VALUE "B".
   88 IS-TYPE-C VALUE "C".
PROCEDURE DIVISION.
    PERFORM PROCESS-PARA.
    STOP RUN.
PROCESS-PARA.
    IF IS-TYPE-A
        DISPLAY "Processing type A"
    ELSE
        IF IS-TYPE-B
            DISPLAY "Processing type B"
        ELSE
            DISPLAY "Processing other type"
        END-IF
    END-IF.
"#);
}

#[test]
 // COMPUTE with subscripted target needs parser work
fn nested_perform_with_compute() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. NESTPERF.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TABLE.
   05 WS-ROW OCCURS 3 TIMES.
      10 WS-COL PIC 9(5) OCCURS 3 TIMES.
01 WS-I PIC 9(3).
01 WS-J PIC 9(3).
PROCEDURE DIVISION.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 3
        PERFORM VARYING WS-J FROM 1 BY 1 UNTIL WS-J > 3
            COMPUTE WS-COL(WS-J) = WS-I * WS-J
        END-PERFORM
    END-PERFORM.
    DISPLAY "Multiplication table done".
    STOP RUN.
"#);
}

#[test]
fn free_allocate_program() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. DYNALLOC.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PTR PIC X(100).
PROCEDURE DIVISION.
    ALLOCATE WS-PTR.
    DISPLAY "Allocated".
    FREE WS-PTR.
    DISPLAY "Freed".
    STOP RUN.
"#);
}
