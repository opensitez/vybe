*> vybe-test: cobol/value_all_forms/value_spaces_and_then_moved_to
*> origin: languages/cobol/tests/cobol/test_value_all_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(8) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "OK" TO S.
    DISPLAY S.
    MOVE SPACES TO WS-VYBE-L
    STRING S DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OK      "
        DISPLAY "FAIL: want [OK      ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

