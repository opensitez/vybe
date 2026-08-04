*> vybe-test: cobol/goto/test_goto_exit_paragraph
*> origin: languages/cobol/tests/cobol/test_goto.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.

    PERFORM MY-PARA.
    STOP RUN.
MY-PARA.
    DISPLAY "START".
    GO TO MY-PARA-EXIT.
    DISPLAY "SKIPPED".
MY-PARA-EXIT.
    EXIT.
    STOP RUN.

