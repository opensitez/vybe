*> vybe-test: cobol/alter_stop/alter_in_loop
*> origin: languages/cobol/tests/cobol/test_alter_stop.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 ws-iter PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           GO TO loop-body.
       loop-body.
           ADD 1 TO ws-iter
           IF ws-iter >= 3
               ALTER loop-exit TO PROCEED TO done
           END-IF
           IF ws-iter < 3
               GO TO loop-body
           ELSE
               GO TO loop-exit
           END-IF.
       loop-exit.
           GO TO loop-body.
       done.
           DISPLAY ws-iter
           STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING ws-iter DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "3"
        DISPLAY "FAIL: want [3] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

