*> vybe-test: cobol/category_string_unstring/test_unstring_delimited_by_all
*> origin: languages/cobol/tests/cobol/test_category_string_unstring.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. UNSTRING-ALL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 SRC PIC X(20) VALUE "A,,,,,B".
       01 OUT-1 PIC X(5).
       01 OUT-2 PIC X(5).
       PROCEDURE DIVISION.
           UNSTRING SRC DELIMITED BY ALL ","
              INTO OUT-1 OUT-2.
           DISPLAY OUT-1 "|" OUT-2.
    MOVE SPACES TO WS-VYBE-L
    STRING OUT-1 DELIMITED SIZE "|" DELIMITED SIZE OUT-2 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A    |B    "
        DISPLAY "FAIL: want [A    |B    ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

