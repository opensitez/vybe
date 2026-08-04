*> vybe-test: cobol/initialize_forms/initialize_group_nested_levels
*> origin: languages/cobol/tests/cobol/test_initialize_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 OUTER.
   05 INNER.
      10 DEEPEST PIC 9(2) VALUE 99.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    INITIALIZE OUTER.
    DISPLAY DEEPEST.
    MOVE SPACES TO WS-VYBE-L
    STRING DEEPEST DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "00"
        DISPLAY "FAIL: want [00] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

