*> vybe-test: cobol/category_string_unstring/test_unstring_multiple_delimiters
*> origin: languages/cobol/tests/cobol/test_category_string_unstring.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. UNSTRING-MULTI.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 SRC PIC X(20) VALUE "A,B;C D".
       01 OUT-1 PIC X(2).
       01 OUT-2 PIC X(2).
       01 OUT-3 PIC X(2).
       01 OUT-4 PIC X(2).
       PROCEDURE DIVISION.
           UNSTRING SRC DELIMITED BY "," OR ";" OR " "
              INTO OUT-1 OUT-2 OUT-3 OUT-4.
           DISPLAY OUT-1 "|" OUT-2 "|" OUT-3 "|" OUT-4.
    MOVE SPACES TO WS-VYBE-L
    STRING OUT-1 DELIMITED SIZE "|" DELIMITED SIZE OUT-2 DELIMITED SIZE "|" DELIMITED SIZE OUT-3 DELIMITED SIZE "|" DELIMITED SIZE OUT-4 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A |B |C |D "
        DISPLAY "FAIL: want [A |B |C |D ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

