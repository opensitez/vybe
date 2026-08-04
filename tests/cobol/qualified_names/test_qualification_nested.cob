*> vybe-test: cobol/qualified_names/test_qualification_nested
*> origin: languages/cobol/tests/cobol/test_qualified_names.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TOP.
   05 WS-SUB.
      10 WS-FIELD PIC X(3) VALUE "XYZ".
PROCEDURE DIVISION.

    DISPLAY WS-FIELD IN WS-SUB IN WS-TOP.
    STOP RUN.

