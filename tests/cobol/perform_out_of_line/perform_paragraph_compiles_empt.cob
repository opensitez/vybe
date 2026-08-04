*> vybe-test: cobol/perform_out_of_line/perform_paragraph_compiles_empty_body
*> origin: languages/cobol/tests/cobol/test_perform_out_of_line.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    PERFORM NOOP.
    STOP RUN.
NOOP.
    CONTINUE.
    STOP RUN.

