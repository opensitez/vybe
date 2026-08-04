*> vybe-test: cobol/paragraph_section_flow/paragraph_exit_before_display
*> origin: languages/cobol/tests/cobol/test_paragraph_section_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 COND PIC 9 VALUE 1.
PROCEDURE DIVISION.
    PERFORM GUARDED.
    STOP RUN.
GUARDED.
    IF COND = 1
        EXIT PARAGRAPH
    END-IF.
    DISPLAY "NOT REACHED".
    STOP RUN.

