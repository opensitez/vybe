*> vybe-test: cobol/accept_forms/stop_run_inside_perform_paragraph
*> origin: languages/cobol/tests/cobol/test_accept_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    PERFORM DO-WORK.
    STOP RUN.
DO-WORK.
    DISPLAY "WORKING".
    STOP RUN.

