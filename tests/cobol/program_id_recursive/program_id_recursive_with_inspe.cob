*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_inspect
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(10) VALUE "HELLO".
01 C PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    INSPECT S TALLYING C FOR ALL "L".
    STOP RUN.

