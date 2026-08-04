*> vybe-test: cobol/program_exit/test_exit_section_paragraph
*> origin: languages/cobol/tests/cobol/test_program_exit.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.

    PERFORM MY-SEC.
    STOP RUN.
MY-SEC SECTION.
MY-PARA.
    DISPLAY "PARA".
    EXIT PARAGRAPH.
    DISPLAY "NOT-SHOWN".
MY-EXIT.
    EXIT SECTION.
    STOP RUN.

