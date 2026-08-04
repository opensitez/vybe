*> vybe-test: cobol/string_and_unstring_extended/string_with_pointer_updates_destination_and_pointer
*> origin: languages/cobol/tests/cobol/test_string_and_unstring_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(2) VALUE "AB".
01 WS-B PIC X(2) VALUE "CD".
01 WS-R PIC X(8) VALUE SPACES.
01 WS-PTR PIC 9(2) VALUE 1.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    STRING WS-A DELIMITED BY SIZE
           WS-B DELIMITED BY SIZE
           INTO WS-R
           WITH POINTER WS-PTR.
    DISPLAY WS-R.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABCD    "
        DISPLAY "FAIL: want [ABCD    ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    DISPLAY WS-PTR.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-PTR DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "05"
        DISPLAY "FAIL: want [05] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

