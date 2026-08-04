*> vybe-test: cobol/go_to_forms/go_to_paragraph_in_section
*> origin: languages/cobol/tests/cobol/test_go_to_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    PERFORM WORK-SEC.
    STOP RUN.
WORK-SEC SECTION.
    GO TO INNER.
    DISPLAY "SKIPPED".
INNER.
    DISPLAY "INNER".
    STOP RUN.

