*> vybe-test: cobol/nested_if_else/if_with_continue_in_branch
*> origin: languages/cobol/tests/cobol/test_nested_if_else.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 1.
PROCEDURE DIVISION.
    IF X = 1
        CONTINUE
    ELSE
        DISPLAY "NO"
    END-IF.
    STOP RUN.

