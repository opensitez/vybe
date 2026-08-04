*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_level88
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 FLAG PIC X VALUE "N".
    88 ENABLED VALUE "Y".
PROCEDURE DIVISION.
    SET ENABLED TO TRUE.
    STOP RUN.

