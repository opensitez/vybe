*> vybe-test: cobol/string_and_unstring_extended/string_with_space_delimiter_concatenates_fields
*> origin: languages/cobol/tests/cobol/test_string_and_unstring_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(4) VALUE "ONE".
01 WS-B PIC X(4) VALUE "TWO".
01 WS-R PIC X(20) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    STRING WS-A DELIMITED BY SPACE
           WS-B DELIMITED BY SPACE
           INTO WS-R.
    DISPLAY WS-R.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ONETWO              "
        DISPLAY "FAIL: want [ONETWO              ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

