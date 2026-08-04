*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_formatted_output
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 FORMATTED PIC ZZ9 VALUE 42.
PROCEDURE DIVISION.
    DISPLAY FORMATTED.
    STOP RUN.

