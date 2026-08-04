*> vybe-test: cobol/accept_environment/accept_environment_missing_var
*> origin: languages/cobol/tests/cobol/test_accept_environment.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 ws-val PIC X(50).
       01 ws-status PIC X(10).
       PROCEDURE DIVISION.
           MOVE SPACES TO ws-val
           ACCEPT ws-val FROM ENVIRONMENT "NONEXISTENT_VAR_XYZ"
           IF ws-val = SPACES
               MOVE "missing" TO ws-status
           ELSE
               MOVE "found" TO ws-status
           END-IF
           DISPLAY ws-status
           STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING ws-status DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "missing"
        DISPLAY "FAIL: want [missing] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

