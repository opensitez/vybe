*> vybe-test: cobol/string_and_unstring_extended/unstring_with_all_delimiter_splits_repeated_separators
*> origin: languages/cobol/tests/cobol/test_string_and_unstring_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SRC PIC X(12) VALUE "A,,B".
01 WS-F1 PIC X(3) VALUE SPACES.
01 WS-F2 PIC X(3) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    UNSTRING WS-SRC DELIMITED BY ALL "," INTO WS-F1 WS-F2.
    DISPLAY WS-F1.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-F1 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A  "
        DISPLAY "FAIL: want [A  ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    DISPLAY WS-F2.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-F2 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "B  "
        DISPLAY "FAIL: want [B  ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

