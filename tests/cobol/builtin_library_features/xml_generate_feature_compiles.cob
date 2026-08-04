*> vybe-test: cobol/builtin_library_features/xml_generate_feature_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_library_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC.
   05 WS-NAME PIC X(10) VALUE "ALICE".
01 WS-XML PIC X(500).
01 WS-LEN PIC 9(5).
PROCEDURE DIVISION.
    XML GENERATE WS-XML FROM WS-REC COUNT IN WS-LEN.
    STOP RUN.

