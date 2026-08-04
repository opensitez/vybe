*> vybe-test: cobol/category_perform/test_perform_thru
*> origin: languages/cobol/tests/cobol/test_category_perform.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PERFORM-THRU.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       PROCEDURE DIVISION.
           PERFORM PARA-A THRU PARA-C.
           DISPLAY "MAIN END".
    MOVE SPACES TO WS-VYBE-L
    STRING "MAIN END" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A"
        DISPLAY "FAIL: want [A] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.
       PARA-A.
           DISPLAY "A".
    MOVE SPACES TO WS-VYBE-L
    STRING "A" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "B"
        DISPLAY "FAIL: want [B] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
       PARA-B.
           DISPLAY "B".
    MOVE SPACES TO WS-VYBE-L
    STRING "B" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "C"
        DISPLAY "FAIL: want [C] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
       PARA-C.
           DISPLAY "C".
    MOVE SPACES TO WS-VYBE-L
    STRING "C" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "MAIN END"
        DISPLAY "FAIL: want [MAIN END] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

