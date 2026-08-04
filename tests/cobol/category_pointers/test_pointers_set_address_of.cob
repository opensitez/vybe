*> vybe-test: cobol/category_pointers/test_pointers_set_address_of
*> origin: languages/cobol/tests/cobol/test_category_pointers.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PTR-ADDRESS.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 WS-VAL PIC X(5) VALUE "HELLO".
       01 WS-PTR USAGE POINTER.
       LINKAGE SECTION.
       01 LK-VAL PIC X(5).
       PROCEDURE DIVISION.
           SET WS-PTR TO ADDRESS OF WS-VAL.
           SET ADDRESS OF LK-VAL TO WS-PTR.
           DISPLAY LK-VAL.
    MOVE SPACES TO WS-VYBE-L
    STRING LK-VAL DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "HELLO"
        DISPLAY "FAIL: want [HELLO] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

