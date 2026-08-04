*> vybe-test: cobol/qualified_names_of_clause/qualified_two_fields_evaluated_together
*> origin: languages/cobol/tests/cobol/test_qualified_names_of_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 HEADER.
   05 TYPE-CODE PIC X VALUE "A".
01 DETAIL.
   05 TYPE-CODE PIC X VALUE "D".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE TYPE-CODE OF HEADER ALSO TYPE-CODE OF DETAIL
        WHEN "A" ALSO "D"
            DISPLAY "VALID"
        WHEN OTHER ALSO OTHER
            DISPLAY "INVALID"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "VALID" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "VALID"
        DISPLAY "FAIL: want [VALID] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

