*> vybe-test: cobol/inspect_converting/inspect_converting_digit_to_asterisk
*> origin: languages/cobol/tests/cobol/test_inspect_converting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(8) VALUE "A1B2C3D4".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    INSPECT S CONVERTING "1234567890" TO "**********".
    DISPLAY S.
    MOVE SPACES TO WS-VYBE-L
    STRING S DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A*B*C*D*"
        DISPLAY "FAIL: want [A*B*C*D*] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

