*> vybe-test: cobol/string_and_unstring_extended/reference_modification_on_string_field
*> origin: languages/cobol/tests/cobol/test_string_and_unstring_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TXT PIC X(10) VALUE "HELLOTEST".
01 WS-SUB PIC X(5) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MOVE WS-TXT(1:5) TO WS-SUB.
    DISPLAY WS-SUB.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-SUB DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "HELLO"
        DISPLAY "FAIL: want [HELLO] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

