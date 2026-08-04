*> vybe-test: cobol/special_names_configuration/special_names_class_digits_runtime
*> origin: languages/cobol/tests/cobol/test_special_names_configuration.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CLASS DIGITS IS "0" THRU "9".
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS PIC X VALUE "7".
PROCEDURE DIVISION.
    IF WS IS DIGITS
        DISPLAY "DIG"
    END-IF
    STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING "DIG" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "DIG"
        DISPLAY "FAIL: want [DIG] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

