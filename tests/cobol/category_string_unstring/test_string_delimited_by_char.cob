*> vybe-test: cobol/category_string_unstring/test_string_delimited_by_char
*> origin: languages/cobol/tests/cobol/test_category_string_unstring.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. STRING-DELIM-CHAR.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 STR-1 PIC X(10) VALUE "HELLO*FOO".
       01 STR-2 PIC X(10) VALUE "WORLD*BAR".
       01 DEST PIC X(20) VALUE SPACES.
       PROCEDURE DIVISION.
           STRING STR-1 DELIMITED BY "*"
                  " " DELIMITED BY SIZE
                  STR-2 DELIMITED BY "*"
                  INTO DEST.
           DISPLAY DEST.
    MOVE SPACES TO WS-VYBE-L
    STRING DEST DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "HELLO WORLD         "
        DISPLAY "FAIL: want [HELLO WORLD         ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

