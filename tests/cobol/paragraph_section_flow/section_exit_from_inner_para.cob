*> vybe-test: cobol/paragraph_section_flow/section_exit_from_inner_para
*> origin: languages/cobol/tests/cobol/test_paragraph_section_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    PERFORM WORK-SEC.
    STOP RUN.
WORK-SEC SECTION.
START-WORK.
    DISPLAY "START".
    EXIT SECTION.
REST-WORK.
    DISPLAY "UNREACHABLE".
    STOP RUN.

