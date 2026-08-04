*> vybe-test: cobol/builtin_library_features/xml_generate_with_attributes_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_library_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC.
   05 WS-ID PIC 9(4) VALUE 1001.
01 WS-XML PIC X(500).
PROCEDURE DIVISION.
    XML GENERATE WS-XML FROM WS-REC WITH ATTRIBUTES.
    STOP RUN.

