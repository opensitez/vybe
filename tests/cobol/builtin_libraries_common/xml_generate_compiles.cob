*> vybe-test: cobol/builtin_libraries_common/xml_generate_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_libraries_common.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC X(10) VALUE "A".
01 X PIC X(200).
01 L PIC 9(5).
PROCEDURE DIVISION.
    XML GENERATE X FROM R COUNT IN L.
    STOP RUN.

