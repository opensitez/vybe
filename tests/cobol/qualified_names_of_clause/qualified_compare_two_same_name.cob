*> vybe-test: cobol/qualified_names_of_clause/qualified_compare_two_same_named_fields
*> origin: languages/cobol/tests/cobol/test_qualified_names_of_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 OLD-REC.
   05 STATUS PIC X VALUE "A".
01 NEW-REC.
   05 STATUS PIC X VALUE "B".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF STATUS OF OLD-REC NOT = STATUS OF NEW-REC
        DISPLAY "CHANGED"
    ELSE
        DISPLAY "SAME"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "CHANGED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "CHANGED"
        DISPLAY "FAIL: want [CHANGED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

