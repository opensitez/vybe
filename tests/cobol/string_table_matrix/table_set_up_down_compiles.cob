*> vybe-test: cobol/string_table_matrix/table_set_up_down_compiles
*> origin: languages/cobol/tests/cobol/test_string_table_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TBL PIC 9 OCCURS 5 TIMES INDEXED BY I.
PROCEDURE DIVISION.
    SET I TO 1.
    SET I UP BY 2.
    SET I DOWN BY 1.
    STOP RUN.

