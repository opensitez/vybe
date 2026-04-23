use super::helpers::compile_ok;

// ── RAISE EXCEPTION with class ────────────────────────────────

#[test] fn raise_exception_class() {
    compile_ok(r#"
       CLASS-ID. AppException INHERITS FROM EXCEPTION-OBJECT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-message PIC X(100).
       METHOD-ID. GET-MESSAGE.
       LINKAGE SECTION.
       01 lk-msg PIC X(100).
       PROCEDURE DIVISION RETURNING lk-msg.
           MOVE ws-message TO lk-msg
           GOBACK.
       END METHOD GET-MESSAGE.
       METHOD-ID. SET-MESSAGE.
       LINKAGE SECTION.
       01 lk-msg PIC X(100).
       PROCEDURE DIVISION USING lk-msg.
           MOVE lk-msg TO ws-message
           GOBACK.
       END METHOD SET-MESSAGE.
       END CLASS AppException.
"#);
}

#[test] fn raise_built_in_exception() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-caught PIC X VALUE "N".
       PROCEDURE DIVISION.
           RAISE EXCEPTION EC-PROGRAM-ARG-OMITTED
           DISPLAY ws-caught
           STOP RUN.
"#);
}

#[test] fn raise_and_resume() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-x PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           ADD 1 TO ws-x
           RAISE EXCEPTION EC-SIZE-OVERFLOW
           ADD 1 TO ws-x
           DISPLAY ws-x
           STOP RUN.
"#);
}

// ── User-defined exception hierarchy ─────────────────────────

#[test] fn custom_exception_class() {
    compile_ok(r#"
       CLASS-ID. ValidationException INHERITS FROM EXCEPTION-OBJECT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-field-name PIC X(30).
       01 ws-error-msg  PIC X(100).
       METHOD-ID. INITIALIZE-EX.
       LINKAGE SECTION.
       01 lk-field PIC X(30).
       01 lk-msg   PIC X(100).
       PROCEDURE DIVISION USING lk-field lk-msg.
           MOVE lk-field TO ws-field-name
           MOVE lk-msg   TO ws-error-msg
           GOBACK.
       END METHOD INITIALIZE-EX.
       END CLASS ValidationException.
"#);
}

#[test] fn exception_hierarchy_two_levels() {
    compile_ok(r#"
       CLASS-ID. BaseException INHERITS FROM EXCEPTION-OBJECT.
       END CLASS BaseException.

       CLASS-ID. DatabaseException INHERITS FROM BaseException.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-sql-code PIC S9(9) COMP.
       METHOD-ID. SET-SQL-CODE.
       LINKAGE SECTION.
       01 lk-code PIC S9(9) COMP.
       PROCEDURE DIVISION USING lk-code.
           MOVE lk-code TO ws-sql-code
           GOBACK.
       END METHOD SET-SQL-CODE.
       END CLASS DatabaseException.
"#);
}

#[test] fn exception_hierarchy_three_levels() {
    compile_ok(r#"
       CLASS-ID. AppBaseEx INHERITS FROM EXCEPTION-OBJECT.
       END CLASS AppBaseEx.

       CLASS-ID. IOException INHERITS FROM AppBaseEx.
       END CLASS IOException.

       CLASS-ID. FileNotFoundException INHERITS FROM IOException.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-filename PIC X(200).
       METHOD-ID. SET-FILENAME.
       LINKAGE SECTION.
       01 lk-fn PIC X(200).
       PROCEDURE DIVISION USING lk-fn.
           MOVE lk-fn TO ws-filename
           GOBACK.
       END METHOD SET-FILENAME.
       END CLASS FileNotFoundException.
"#);
}

// ── INVOKE on exception objects ───────────────────────────────

#[test] fn invoke_exception_method() {
    compile_ok(r#"
       CLASS-ID. MyException INHERITS FROM EXCEPTION-OBJECT.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-code PIC 9(4).
       METHOD-ID. GET-CODE.
       LINKAGE SECTION.
       01 lk-code PIC 9(4).
       PROCEDURE DIVISION RETURNING lk-code.
           MOVE ws-code TO lk-code
           GOBACK.
       END METHOD GET-CODE.
       END CLASS MyException.

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS MyException AS "MyException".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-ex   OBJECT REFERENCE MyException.
       01 ws-code PIC 9(4).
       PROCEDURE DIVISION.
           INVOKE MyException NEW RETURNING ws-ex
           INVOKE ws-ex GET-CODE RETURNING ws-code
           DISPLAY ws-code
           STOP RUN.
"#);
}

// ── Standard exception conditions ────────────────────────────

#[test] fn ec_program_arg_omitted() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-err PIC X VALUE "N".
       PROCEDURE DIVISION.
           IF ws-err = "N"
               RAISE EXCEPTION EC-PROGRAM-ARG-OMITTED
           END-IF
           STOP RUN.
"#);
}

#[test] fn ec_size_overflow() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-n PIC 9 VALUE 9.
       PROCEDURE DIVISION.
           ADD 5 TO ws-n
               ON SIZE ERROR RAISE EXCEPTION EC-SIZE-OVERFLOW
           END-ADD
           STOP RUN.
"#);
}

#[test] fn ec_bound_ptr_null() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-ptr USAGE POINTER VALUE NULL.
       PROCEDURE DIVISION.
           IF ws-ptr = NULL
               RAISE EXCEPTION EC-BOUND-PTR-NULL
           END-IF
           DISPLAY "checked"
           STOP RUN.
"#);
}

#[test] fn ec_io_file_missing() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT data-file ASSIGN TO "nonexistent.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD data-file.
       01 data-rec PIC X(80).
       WORKING-STORAGE SECTION.
       01 ws-handled PIC X VALUE "N".
       PROCEDURE DIVISION.
       DECLARATIVES.
       file-err SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON data-file.
           MOVE "Y" TO ws-handled.
       END DECLARATIVES.
       main-para SECTION.
           OPEN INPUT data-file
           IF ws-handled = "N"
               DISPLAY "opened ok"
           ELSE
               DISPLAY "file error handled"
           END-IF
           STOP RUN.
"#);
}

// ── Exception propagation ─────────────────────────────────────

#[test] fn exception_in_subprogram() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. main-prog.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC X VALUE "N".
       PROCEDURE DIVISION.
           CALL "sub-prog" USING ws-result
               ON EXCEPTION MOVE "E" TO ws-result
           END-CALL
           DISPLAY ws-result
           STOP RUN.

       IDENTIFICATION DIVISION.
       PROGRAM-ID. sub-prog.
       DATA DIVISION.
       LINKAGE SECTION.
       01 lk-result PIC X.
       PROCEDURE DIVISION USING lk-result.
           MOVE "Y" TO lk-result
           GOBACK.
"#);
}

#[test] fn exception_object_null_reference() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       CONFIGURATION SECTION.
       REPOSITORY.
           CLASS EXCEPTION-OBJECT AS "EXCEPTION-OBJECT".
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-obj OBJECT REFERENCE EXCEPTION-OBJECT VALUE NULL.
       PROCEDURE DIVISION.
           IF ws-obj = NULL
               DISPLAY "null object reference"
           END-IF
           STOP RUN.
"#);
}

#[test] fn resume_after_exception() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-step  PIC 9 VALUE 0.
       01 ws-total PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           MOVE 1 TO ws-step
           ADD ws-step TO ws-total
           RAISE EXCEPTION EC-PROGRAM-RECURSIVE-CALL
           MOVE 2 TO ws-step
           ADD ws-step TO ws-total
           DISPLAY ws-total
           STOP RUN.
"#);
}
