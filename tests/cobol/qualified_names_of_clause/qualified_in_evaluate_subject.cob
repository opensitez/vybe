*> vybe-test: cobol/qualified_names_of_clause/qualified_in_evaluate_subject
*> origin: languages/cobol/tests/cobol/test_qualified_names_of_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 REC-X.
   05 TYPE PIC X VALUE "A".
01 REC-Y.
   05 TYPE PIC X VALUE "B".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE TYPE OF REC-X
        WHEN "A"
            DISPLAY "TYPE A"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "TYPE A" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "TYPE A"
        DISPLAY "FAIL: want [TYPE A] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

