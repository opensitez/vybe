*> vybe-test: cobol/string_functions_intrinsic/intrinsic_formatted_datetime_compiles
*> origin: languages/cobol/tests/cobol/test_string_functions_intrinsic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 DT PIC X(30).
PROCEDURE DIVISION.
    MOVE FUNCTION FORMATTED-DATETIME("YYYY-MM-DDThh:mm:ss" 20230615 1430) TO DT.
    STOP RUN.

