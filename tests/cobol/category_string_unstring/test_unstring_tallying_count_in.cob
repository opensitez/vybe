*> vybe-test: cobol/category_string_unstring/test_unstring_tallying_count_in
*> origin: languages/cobol/tests/cobol/test_category_string_unstring.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. UNSTRING-TALLY.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 SRC PIC X(20) VALUE "APPLE,BANANA".
       01 OUT-1 PIC X(10).
       01 CNT-1 PIC 9(2) VALUE 0.
       01 OUT-2 PIC X(10).
       01 CNT-2 PIC 9(2) VALUE 0.
       01 TALLY PIC 9(2) VALUE 0.
       PROCEDURE DIVISION.
           UNSTRING SRC DELIMITED BY ","
              INTO OUT-1 COUNT IN CNT-1
                   OUT-2 COUNT IN CNT-2
              TALLYING IN TALLY.
           DISPLAY CNT-1 " " CNT-2 " " TALLY.
    MOVE SPACES TO WS-VYBE-L
    STRING CNT-1 DELIMITED SIZE " " DELIMITED SIZE CNT-2 DELIMITED SIZE " " DELIMITED SIZE TALLY DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "05 06 02"
        DISPLAY "FAIL: want [05 06 02] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

