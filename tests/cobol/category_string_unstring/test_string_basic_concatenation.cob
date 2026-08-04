*> vybe-test: cobol/category_string_unstring/test_string_basic_concatenation
*> origin: languages/cobol/tests/cobol/test_category_string_unstring.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. STRING-BASIC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 STR-1 PIC X(3) VALUE "ABC".
       01 STR-2 PIC X(3) VALUE "DEF".
       01 DEST PIC X(10) VALUE SPACES.
       PROCEDURE DIVISION.
           STRING STR-1 DELIMITED BY SIZE
                  STR-2 DELIMITED BY SIZE
                  INTO DEST.
           DISPLAY DEST.
    MOVE SPACES TO WS-VYBE-L
    STRING DEST DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABCDEF    "
        DISPLAY "FAIL: want [ABCDEF    ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

