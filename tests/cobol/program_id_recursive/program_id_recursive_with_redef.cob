*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_redefines
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 UNION-FIELD PIC X(4) VALUE "ABCD".
01 UNION-NUM REDEFINES UNION-FIELD PIC 9(4).
PROCEDURE DIVISION.
    DISPLAY UNION-NUM.
    STOP RUN.

