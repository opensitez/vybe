*> vybe-test: cobol/paragraph_section_flow/section_exit_terminates_section
*> origin: languages/cobol/tests/cobol/test_paragraph_section_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    PERFORM MY-SECT.
    STOP RUN.
MY-SECT SECTION.
    DISPLAY "START".
    EXIT SECTION.
    DISPLAY "UNREACHABLE".
    STOP RUN.

