*> vybe-test: cobol/category_procedure_division_sort_statement/sort_statement_with_input_and_output_procedures_runtime
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_sort_statement.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. SRT1.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT S ASSIGN TO "SRT1".
DATA DIVISION.
FILE SECTION.
SD S.
01 R.
    05 K PIC X(1).
    05 V PIC X(3).
PROCEDURE DIVISION.
    SORT S
        ON ASCENDING KEY K
        INPUT PROCEDURE IS SRT-IN
        OUTPUT PROCEDURE IS SRT-OUT.
    STOP RUN.
SRT-IN SECTION.
    MOVE "C" TO R
    RELEASE R
    MOVE "A" TO R
    RELEASE R
    MOVE "B" TO R
    RELEASE R.
SRT-OUT SECTION.
    RETURN S AT END DISPLAY "DONE" END-RETURN
    PERFORM UNTIL FALSE
        RETURN S
            AT END DISPLAY "DONE"
            GO TO SRT-OUT-DONE
            NOT AT END DISPLAY K
        END-RETURN
    END-PERFORM
SRT-OUT-DONE.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING "DONE" DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "A"
                DISPLAY "FAIL at 1 want [A] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "B"
                DISPLAY "FAIL at 2 want [B] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 3
            IF WS-VYBE-L NOT = "C"
                DISPLAY "FAIL at 3 want [C] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 4
            IF WS-VYBE-L NOT = "DONE"
                DISPLAY "FAIL at 4 want [DONE] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 4 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.
    EXIT.

    IF WS-VYBE-I NOT = 4
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 4"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

