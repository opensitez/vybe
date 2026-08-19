*> vybe-test: cobol/condition_names/condition_name_supports_multiple_values_with_false_item
*> origin: languages/cobol/tests/cobol/test_condition_names.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-CODE PIC X VALUE "D".
   88 IS-VALID VALUE "A", "B", "C".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    IF IS-VALID
        DISPLAY "YES"
    ELSE
        DISPLAY "NO"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "YES" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "YES"
        DISPLAY "FAIL: want [YES] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

