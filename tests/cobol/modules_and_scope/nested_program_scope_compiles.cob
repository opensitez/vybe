*> vybe-test: cobol/modules_and_scope/nested_program_scope_compiles
*> origin: languages/cobol/tests/cobol/test_modules_and_scope.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. PARENT.
PROCEDURE DIVISION.
    DISPLAY "PARENT".
    STOP RUN.
END PROGRAM PARENT.
IDENTIFICATION DIVISION.
PROGRAM-ID. CHILD.
PROCEDURE DIVISION.
    DISPLAY "CHILD".
    STOP RUN.
END PROGRAM CHILD.

