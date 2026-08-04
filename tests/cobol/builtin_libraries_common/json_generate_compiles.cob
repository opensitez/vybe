*> vybe-test: cobol/builtin_libraries_common/json_generate_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_libraries_common.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R.
   05 N PIC X(10) VALUE "A".
01 J PIC X(100).
PROCEDURE DIVISION.
    JSON GENERATE J FROM R.
    STOP RUN.

