*> vybe-test: cobol/qualified_names_of_clause/qualified_level88_in_qualified_field
*> origin: languages/cobol/tests/cobol/test_qualified_names_of_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 FLAG-REC.
   05 ACTIVE-FLAG PIC X VALUE "Y".
      88 ACTIVE-ON VALUE "Y".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF ACTIVE-ON
        DISPLAY "ACTIVE"
    ELSE
        DISPLAY "INACTIVE"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "ACTIVE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ACTIVE"
        DISPLAY "FAIL: want [ACTIVE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

