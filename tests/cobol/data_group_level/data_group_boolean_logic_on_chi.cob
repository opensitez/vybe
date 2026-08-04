*> vybe-test: cobol/data_group_level/data_group_boolean_logic_on_child
*> origin: languages/cobol/tests/cobol/test_data_group_level.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 STATUS-REC.
   05 STATE-CODE PIC X VALUE "A".
   05 STATE-NUM PIC 9 VALUE 1.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF STATE-CODE = "A" AND STATE-NUM = 1
        DISPLAY "VALID"
    ELSE
        DISPLAY "INVALID"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "VALID" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "VALID"
        DISPLAY "FAIL: want [VALID] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

