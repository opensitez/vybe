*> vybe-test: cobol/cobol2023_nested/nested_program_basic
*> origin: languages/cobol/tests/cobol/test_cobol2023_nested.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-PROG.
PROCEDURE DIVISION.
    DISPLAY "Main program".
    STOP RUN.
IDENTIFICATION DIVISION.
PROGRAM-ID. HELPER.
PROCEDURE DIVISION.
    DISPLAY "Helper".
    STOP RUN.
END PROGRAM HELPER.
END PROGRAM MAIN-PROG.

