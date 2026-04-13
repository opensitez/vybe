use super::helpers::{compile_ok, parse_ok, compile_ok_check};




// ═══════════════════════════════════════════════════════════
// PROGRAM STRUCTURE
// ═══════════════════════════════════════════════════════════
#[test]
fn minimal_program() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. HELLO.
PROCEDURE DIVISION.
    DISPLAY "Hello, World!".
    STOP RUN.
"#);
}

#[test]
fn program_with_data() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. VARS.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(20) VALUE "Alice".
01 WS-AGE  PIC 9(3)  VALUE 30.
PROCEDURE DIVISION.
    DISPLAY WS-NAME.
    DISPLAY WS-AGE.
    STOP RUN.
"#);
}

#[test]
fn program_with_author() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. META.
AUTHOR. Test Author.
DATE-WRITTEN. 2024-01-01.
PROCEDURE DIVISION.
    DISPLAY "Hello".
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// DATA DIVISION — LEVEL NUMBERS & PIC
// ═══════════════════════════════════════════════════════════
#[test]
fn pic_alpha() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. PICALPHA.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(30) VALUE "Hello World".
PROCEDURE DIVISION.
    DISPLAY WS-TEXT.
    STOP RUN.
"#);
}

#[test]
fn pic_numeric() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. PICNUM.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-COUNT  PIC 9(5)    VALUE 12345.
01 WS-AMOUNT PIC 9(7)V99 VALUE 1234.56.
01 WS-SIGNED PIC S9(5)   VALUE -100.
PROCEDURE DIVISION.
    DISPLAY WS-COUNT.
    DISPLAY WS-AMOUNT.
    DISPLAY WS-SIGNED.
    STOP RUN.
"#);
}

#[test]
fn group_items() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. GROUPS.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PERSON.
   05 WS-FIRST-NAME PIC X(15) VALUE "John".
   05 WS-LAST-NAME  PIC X(15) VALUE "Doe".
   05 WS-AGE        PIC 9(3)  VALUE 25.
PROCEDURE DIVISION.
    DISPLAY WS-FIRST-NAME.
    DISPLAY WS-LAST-NAME.
    DISPLAY WS-AGE.
    STOP RUN.
"#);
}

#[test]
fn level_88_conditions() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. COND88.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STATUS PIC X(1) VALUE "A".
   88 IS-ACTIVE  VALUE "A".
   88 IS-INACTIVE VALUE "I".
PROCEDURE DIVISION.
    IF IS-ACTIVE
        DISPLAY "Active"
    END-IF.
    STOP RUN.
"#);
}

#[test]
fn occurs_table() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. TABLES.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TABLE.
   05 WS-ITEM PIC X(10) OCCURS 5 TIMES.
PROCEDURE DIVISION.
    MOVE "First"  TO WS-ITEM(1).
    MOVE "Second" TO WS-ITEM(2).
    DISPLAY WS-ITEM(1).
    DISPLAY WS-ITEM(2).
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// ARITHMETIC
// ═══════════════════════════════════════════════════════════
#[test]
fn add_statement() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. ARITH.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(5) VALUE 10.
01 WS-B PIC 9(5) VALUE 20.
01 WS-C PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    ADD WS-A TO WS-B.
    ADD WS-A WS-B GIVING WS-C.
    STOP RUN.
"#);
}

#[test]
fn subtract_statement() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. SUB.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(5) VALUE 50.
01 WS-B PIC 9(5) VALUE 20.
01 WS-C PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    SUBTRACT WS-B FROM WS-A.
    SUBTRACT WS-B FROM WS-A GIVING WS-C.
    STOP RUN.
"#);
}

#[test]
fn multiply_statement() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. MUL.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(5) VALUE 5.
01 WS-B PIC 9(5) VALUE 3.
01 WS-C PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    MULTIPLY WS-A BY WS-B.
    MULTIPLY WS-A BY WS-B GIVING WS-C.
    STOP RUN.
"#);
}

#[test]
fn divide_statement() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. DIV.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(5) VALUE 100.
01 WS-B PIC 9(5) VALUE 3.
01 WS-C PIC 9(5) VALUE 0.
01 WS-R PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    DIVIDE WS-A BY WS-B GIVING WS-C REMAINDER WS-R.
    STOP RUN.
"#);
}

#[test]
fn compute_statement() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. COMP.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(5) VALUE 10.
01 WS-B PIC 9(5) VALUE 3.
01 WS-RESULT PIC 9(10) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE WS-RESULT = WS-A + WS-B * 2.
    COMPUTE WS-RESULT = (WS-A + WS-B) * 2.
    COMPUTE WS-RESULT = WS-A ** 2.
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// MOVE STATEMENT
// ═══════════════════════════════════════════════════════════
#[test]
fn move_basic() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. MOV.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(10) VALUE SPACES.
01 WS-B PIC 9(5)  VALUE 0.
PROCEDURE DIVISION.
    MOVE "Hello" TO WS-A.
    MOVE 42 TO WS-B.
    STOP RUN.
"#);
}

#[test]
fn move_corresponding() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. MOVCORR.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SRC.
   05 WS-NAME PIC X(10) VALUE "Alice".
   05 WS-AGE  PIC 9(3)  VALUE 30.
01 WS-DST.
   05 WS-NAME PIC X(10).
   05 WS-AGE  PIC 9(3).
PROCEDURE DIVISION.
    MOVE CORRESPONDING WS-SRC TO WS-DST.
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// CONTROL FLOW
// ═══════════════════════════════════════════════════════════
#[test]
fn if_else() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. IFELSE.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-AGE PIC 9(3) VALUE 25.
PROCEDURE DIVISION.
    IF WS-AGE >= 18
        DISPLAY "Adult"
    ELSE
        DISPLAY "Minor"
    END-IF.
    STOP RUN.
"#);
}

#[test]
fn if_nested() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. IFNEST.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SCORE PIC 9(3) VALUE 85.
PROCEDURE DIVISION.
    IF WS-SCORE >= 90
        DISPLAY "A"
    ELSE
        IF WS-SCORE >= 80
            DISPLAY "B"
        ELSE
            IF WS-SCORE >= 70
                DISPLAY "C"
            ELSE
                DISPLAY "F"
            END-IF
        END-IF
    END-IF.
    STOP RUN.
"#);
}

#[test]
fn evaluate_when() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. EVAL.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-GRADE PIC X(1) VALUE "B".
PROCEDURE DIVISION.
    EVALUATE WS-GRADE
        WHEN "A"
            DISPLAY "Excellent"
        WHEN "B"
            DISPLAY "Good"
        WHEN "C"
            DISPLAY "Average"
        WHEN OTHER
            DISPLAY "Unknown"
    END-EVALUATE.
    STOP RUN.
"#);
}

#[test]
fn evaluate_true() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. EVTRUE.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEMP PIC S9(3) VALUE 25.
PROCEDURE DIVISION.
    EVALUATE TRUE
        WHEN WS-TEMP > 30
            DISPLAY "Hot"
        WHEN WS-TEMP > 20
            DISPLAY "Warm"
        WHEN WS-TEMP > 10
            DISPLAY "Cool"
        WHEN OTHER
            DISPLAY "Cold"
    END-EVALUATE.
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// PERFORM (LOOPS)
// ═══════════════════════════════════════════════════════════
#[test]
fn perform_times() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. PTIMES.
PROCEDURE DIVISION.
    PERFORM 5 TIMES
        DISPLAY "Hello"
    END-PERFORM.
    STOP RUN.
"#);
}

#[test]
fn perform_until() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. PUNTIL.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-I PIC 9(3) VALUE 1.
PROCEDURE DIVISION.
    PERFORM UNTIL WS-I > 10
        DISPLAY WS-I
        ADD 1 TO WS-I
    END-PERFORM.
    STOP RUN.
"#);
}

#[test]
fn perform_varying() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. PVARY.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-I PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 10
        DISPLAY WS-I
    END-PERFORM.
    STOP RUN.
"#);
}

#[test]
fn perform_paragraph() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. PPARA.
PROCEDURE DIVISION.
    PERFORM GREET-PARA.
    STOP RUN.
GREET-PARA.
    DISPLAY "Hello from paragraph".
"#);
}

// ═══════════════════════════════════════════════════════════
// STRING OPERATIONS
// ═══════════════════════════════════════════════════════════
#[test]
fn string_concat() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. STRCAT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-FIRST  PIC X(10) VALUE "Hello".
01 WS-SECOND PIC X(10) VALUE "World".
01 WS-RESULT PIC X(25) VALUE SPACES.
PROCEDURE DIVISION.
    STRING WS-FIRST DELIMITED BY SPACE
           " "      DELIMITED BY SIZE
           WS-SECOND DELIMITED BY SPACE
           INTO WS-RESULT.
    DISPLAY WS-RESULT.
    STOP RUN.
"#);
}

#[test]
fn unstring_split() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. UNSPLIT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-INPUT  PIC X(30) VALUE "John,Doe,30".
01 WS-FIRST  PIC X(10).
01 WS-LAST   PIC X(10).
01 WS-AGE    PIC X(5).
PROCEDURE DIVISION.
    UNSTRING WS-INPUT DELIMITED BY ","
        INTO WS-FIRST WS-LAST WS-AGE.
    DISPLAY WS-FIRST.
    DISPLAY WS-LAST.
    DISPLAY WS-AGE.
    STOP RUN.
"#);
}

#[test]
fn inspect_tallying() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. INSP.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT  PIC X(20) VALUE "Hello World".
01 WS-COUNT PIC 9(3)  VALUE 0.
PROCEDURE DIVISION.
    INSPECT WS-TEXT TALLYING WS-COUNT FOR ALL "l".
    DISPLAY WS-COUNT.
    STOP RUN.
"#);
}

#[test]
fn inspect_replacing() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. INSPR.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(20) VALUE "Hello World".
PROCEDURE DIVISION.
    INSPECT WS-TEXT REPLACING ALL "l" BY "r".
    DISPLAY WS-TEXT.
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// INTRINSIC FUNCTIONS (COBOL 2023)
// ═══════════════════════════════════════════════════════════
#[test]
fn func_length() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. FLEN.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(20) VALUE "Hello".
01 WS-LEN  PIC 9(5)  VALUE 0.
PROCEDURE DIVISION.
    MOVE FUNCTION LENGTH(WS-TEXT) TO WS-LEN.
    DISPLAY WS-LEN.
    STOP RUN.
"#);
}

#[test]
fn func_upper_lower() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. FCASE.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(20) VALUE "Hello World".
01 WS-UP   PIC X(20).
01 WS-LOW  PIC X(20).
PROCEDURE DIVISION.
    MOVE FUNCTION UPPER-CASE(WS-TEXT) TO WS-UP.
    MOVE FUNCTION LOWER-CASE(WS-TEXT) TO WS-LOW.
    DISPLAY WS-UP.
    DISPLAY WS-LOW.
    STOP RUN.
"#);
}

#[test]
fn func_trim() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. FTRIM.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(20) VALUE "  Hello  ".
01 WS-OUT  PIC X(20).
PROCEDURE DIVISION.
    MOVE FUNCTION TRIM(WS-TEXT) TO WS-OUT.
    DISPLAY WS-OUT.
    STOP RUN.
"#);
}

#[test]
fn func_reverse() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. FREV.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(10) VALUE "Hello".
01 WS-OUT  PIC X(10).
PROCEDURE DIVISION.
    MOVE FUNCTION REVERSE(WS-TEXT) TO WS-OUT.
    DISPLAY WS-OUT.
    STOP RUN.
"#);
}

#[test]
fn func_current_date() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. FDATE.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DATE PIC X(21).
PROCEDURE DIVISION.
    MOVE FUNCTION CURRENT-DATE TO WS-DATE.
    DISPLAY WS-DATE.
    STOP RUN.
"#);
}

#[test]
fn func_max_min() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. FMINMAX.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RESULT PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    MOVE FUNCTION MAX(10 20 30) TO WS-RESULT.
    DISPLAY WS-RESULT.
    MOVE FUNCTION MIN(10 20 30) TO WS-RESULT.
    DISPLAY WS-RESULT.
    STOP RUN.
"#);
}

#[test]
fn func_mod_rem() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. FMOD.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RESULT PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    MOVE FUNCTION MOD(17 5) TO WS-RESULT.
    DISPLAY WS-RESULT.
    STOP RUN.
"#);
}

#[test]
fn func_numval() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. FNUMVAL.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT   PIC X(10) VALUE "12345".
01 WS-NUMBER PIC 9(10) VALUE 0.
PROCEDURE DIVISION.
    MOVE FUNCTION NUMVAL(WS-TEXT) TO WS-NUMBER.
    DISPLAY WS-NUMBER.
    STOP RUN.
"#);
}

#[test]
fn func_substitute() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. FSUB.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(30) VALUE "Hello World".
01 WS-OUT  PIC X(30).
PROCEDURE DIVISION.
    MOVE FUNCTION SUBSTITUTE(WS-TEXT "World" "COBOL")
         TO WS-OUT.
    DISPLAY WS-OUT.
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// FILE I/O
// ═══════════════════════════════════════════════════════════
#[test]
fn file_read_write() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. FILEIO.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RECORD PIC X(80).
PROCEDURE DIVISION.
    DISPLAY "File I/O test".
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// ACCEPT STATEMENT
// ═══════════════════════════════════════════════════════════
#[test]
fn accept_input() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. INPUT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(20).
PROCEDURE DIVISION.
    DISPLAY "Enter name: ".
    ACCEPT WS-NAME.
    DISPLAY "Hello " WS-NAME.
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// INITIALIZE
// ═══════════════════════════════════════════════════════════
#[test]
fn initialize_stmt() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. INIT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC.
   05 WS-NAME PIC X(10) VALUE "Old".
   05 WS-AGE  PIC 9(3)  VALUE 99.
PROCEDURE DIVISION.
    INITIALIZE WS-REC.
    DISPLAY WS-NAME.
    DISPLAY WS-AGE.
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// JSON (COBOL 2023)
// ═══════════════════════════════════════════════════════════
#[test]
fn json_generate() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. JSONGEN.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PERSON.
   05 WS-NAME PIC X(10) VALUE "Alice".
   05 WS-AGE  PIC 9(3)  VALUE 30.
01 WS-JSON PIC X(100).
PROCEDURE DIVISION.
    JSON GENERATE WS-JSON FROM WS-PERSON.
    DISPLAY WS-JSON.
    STOP RUN.
"#);
}

#[test]
fn json_parse() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. JSONPAR.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-JSON PIC X(100) VALUE '{"name":"Alice","age":30}'.
01 WS-PERSON.
   05 WS-NAME PIC X(10).
   05 WS-AGE  PIC 9(3).
PROCEDURE DIVISION.
    JSON PARSE WS-JSON INTO WS-PERSON.
    DISPLAY WS-NAME.
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// EXCEPTION HANDLING (COBOL 2023)
// ═══════════════════════════════════════════════════════════
#[test]
fn raise_exception() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. EXCEPT.
PROCEDURE DIVISION.
    RAISE EXCEPTION "Something went wrong".
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// GOBACK
// ═══════════════════════════════════════════════════════════
#[test]
fn goback_stmt() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. GOBACK1.
PROCEDURE DIVISION.
    DISPLAY "Done".
    GOBACK.
"#);
}

// ═══════════════════════════════════════════════════════════
// CALL STATEMENT
// ═══════════════════════════════════════════════════════════
#[test]
fn call_subprogram() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. CALLER.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RESULT PIC 9(5).
PROCEDURE DIVISION.
    CALL "SUBPROG" USING WS-RESULT.
    DISPLAY WS-RESULT.
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// SEARCH (table lookup)
// ═══════════════════════════════════════════════════════════
#[test]
fn search_table() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. SRCH.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TABLE.
   05 WS-ENTRY OCCURS 10 TIMES.
      10 WS-KEY   PIC 9(3).
      10 WS-VALUE PIC X(10).
01 WS-IDX PIC 9(3) VALUE 1.
PROCEDURE DIVISION.
    DISPLAY "Search test".
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// SET STATEMENT
// ═══════════════════════════════════════════════════════════
#[test]
fn set_statement() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. SETST.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-FLAG PIC 9(1) VALUE 0.
   88 IS-ON  VALUE 1.
   88 IS-OFF VALUE 0.
PROCEDURE DIVISION.
    SET IS-ON TO TRUE.
    DISPLAY WS-FLAG.
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// CONTINUE
// ═══════════════════════════════════════════════════════════
#[test]
fn continue_stmt() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. CONT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC 9(3) VALUE 5.
PROCEDURE DIVISION.
    IF WS-X > 10
        DISPLAY "Big"
    ELSE
        CONTINUE
    END-IF.
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// COMPLEX PROGRAMS
// ═══════════════════════════════════════════════════════════
#[test]
fn fizzbuzz() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. FIZZBUZZ.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-I      PIC 9(3) VALUE 0.
01 WS-MOD3   PIC 9(3) VALUE 0.
01 WS-MOD5   PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 15
        COMPUTE WS-MOD3 = FUNCTION MOD(WS-I 3)
        COMPUTE WS-MOD5 = FUNCTION MOD(WS-I 5)
        EVALUATE TRUE
            WHEN WS-MOD3 = 0 AND WS-MOD5 = 0
                DISPLAY "FizzBuzz"
            WHEN WS-MOD3 = 0
                DISPLAY "Fizz"
            WHEN WS-MOD5 = 0
                DISPLAY "Buzz"
            WHEN OTHER
                DISPLAY WS-I
        END-EVALUATE
    END-PERFORM.
    STOP RUN.
"#);
}

#[test]
fn factorial() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. FACTORIAL.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-N      PIC 9(3)  VALUE 10.
01 WS-RESULT PIC 9(15) VALUE 1.
01 WS-I      PIC 9(3)  VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-N
        MULTIPLY WS-I BY WS-RESULT
    END-PERFORM.
    DISPLAY "Factorial of " WS-N " = " WS-RESULT.
    STOP RUN.
"#);
}

#[test]
fn temperature_converter() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. TEMPCONV.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-CELSIUS    PIC S9(5)V99 VALUE 100.
01 WS-FAHRENHEIT PIC S9(5)V99 VALUE 0.
PROCEDURE DIVISION.
    COMPUTE WS-FAHRENHEIT = (WS-CELSIUS * 9 / 5) + 32.
    DISPLAY "Celsius: " WS-CELSIUS.
    DISPLAY "Fahrenheit: " WS-FAHRENHEIT.
    STOP RUN.
"#);
}

#[test]
fn string_processing() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. STRPROC.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-INPUT  PIC X(50) VALUE "  Hello, COBOL World!  ".
01 WS-TRIMMED PIC X(50).
01 WS-UPPER   PIC X(50).
01 WS-LOWER   PIC X(50).
01 WS-LEN     PIC 9(5).
PROCEDURE DIVISION.
    MOVE FUNCTION TRIM(WS-INPUT) TO WS-TRIMMED.
    MOVE FUNCTION UPPER-CASE(WS-TRIMMED) TO WS-UPPER.
    MOVE FUNCTION LOWER-CASE(WS-TRIMMED) TO WS-LOWER.
    MOVE FUNCTION LENGTH(WS-TRIMMED) TO WS-LEN.
    DISPLAY "Trimmed: " WS-TRIMMED.
    DISPLAY "Upper: " WS-UPPER.
    DISPLAY "Lower: " WS-LOWER.
    DISPLAY "Length: " WS-LEN.
    STOP RUN.
"#);
}

#[test]
fn paragraph_perform_thru() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. PARATHRU.
PROCEDURE DIVISION.
    PERFORM INIT-PARA.
    PERFORM PROCESS-PARA.
    STOP RUN.
INIT-PARA.
    DISPLAY "Initializing".
PROCESS-PARA.
    DISPLAY "Processing".
"#);
}

#[test]
fn multiple_conditions() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. MULTICOND.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 5.
01 WS-B PIC 9(3) VALUE 10.
01 WS-C PIC 9(3) VALUE 15.
PROCEDURE DIVISION.
    IF WS-A < WS-B AND WS-B < WS-C
        DISPLAY "Ascending"
    END-IF.
    IF WS-A = 5 OR WS-B = 5
        DISPLAY "One is five"
    END-IF.
    IF NOT WS-A = 0
        DISPLAY "A is not zero"
    END-IF.
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// COBOL 2023 SPECIFIC
// ═══════════════════════════════════════════════════════════
#[test]
fn boolean_type() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. BOOL.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-FLAG PIC 9(1) VALUE 1.
PROCEDURE DIVISION.
    IF WS-FLAG = 1
        DISPLAY "True"
    ELSE
        DISPLAY "False"
    END-IF.
    STOP RUN.
"#);
}

#[test]
fn display_multiple() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. DISPMUL.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(10) VALUE "Alice".
01 WS-AGE  PIC 9(3)  VALUE 30.
PROCEDURE DIVISION.
    DISPLAY "Name: " WS-NAME " Age: " WS-AGE.
    STOP RUN.
"#);
}
