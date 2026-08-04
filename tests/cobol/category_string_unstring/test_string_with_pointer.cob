*> vybe-test: cobol/category_string_unstring/test_string_with_pointer
*> origin: languages/cobol/tests/cobol/test_category_string_unstring.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. STRING-PTR.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 STR-1 PIC X(5) VALUE "COBOL".
       01 DEST PIC X(10) VALUE SPACES.
       01 PTR PIC 9(2) VALUE 3.
       PROCEDURE DIVISION.
           STRING STR-1 DELIMITED BY SIZE
                  INTO DEST WITH POINTER PTR.
           DISPLAY DEST.
    MOVE SPACES TO WS-VYBE-L
    STRING DEST DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "  COBOL   "
        DISPLAY "FAIL: want [  COBOL   ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           DISPLAY PTR.
    MOVE SPACES TO WS-VYBE-L
    STRING PTR DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "08"
        DISPLAY "FAIL: want [08] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

