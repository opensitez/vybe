*> vybe-test: cobol/qualified_names/test_qualification_nested_in_section
*> origin: languages/cobol/tests/cobol/test_qualified_names.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-OUTER.
   05 WS-LEFT PIC X(5) VALUE "LEFT ".
   05 WS-INNER.
      10 WS-ALIAS PIC X(5) VALUE "ALIAS".
PROCEDURE DIVISION.

    DISPLAY WS-ALIAS IN WS-INNER IN WS-OUTER.
    DISPLAY WS-LEFT IN WS-OUTER.
    MOVE "NEW" TO WS-ALIAS IN WS-INNER IN WS-OUTER.
    STOP RUN.

