*> vybe-test: cobol/paragraph_section_flow/paragraph_exit_paragraph_compiles
*> origin: languages/cobol/tests/cobol/test_paragraph_section_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 FLAG PIC 9 VALUE 1.
PROCEDURE DIVISION.
    PERFORM DO-STUFF.
    STOP RUN.
DO-STUFF.
    IF FLAG = 0
        EXIT PARAGRAPH
    END-IF.
    DISPLAY "CONTINUING".
    STOP RUN.

