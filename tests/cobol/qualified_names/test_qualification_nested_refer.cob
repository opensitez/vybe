*> vybe-test: cobol/qualified_names/test_qualification_nested_reference_in_expression
*> origin: languages/cobol/tests/cobol/test_qualified_names.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TOP.
   05 WS-V1 PIC 9(2) VALUE 10.
   05 WS-CHILD.
      10 WS-V2 PIC 9(2) VALUE 20.
PROCEDURE DIVISION.

    ADD 1 TO WS-V2 IN WS-CHILD IN WS-TOP
    IF WS-V2 IN WS-CHILD IN WS-TOP > WS-V1 IN WS-TOP
        MOVE WS-V1 IN WS-TOP TO WS-V2 IN WS-CHILD IN WS-TOP
    END-IF
    STOP RUN.

