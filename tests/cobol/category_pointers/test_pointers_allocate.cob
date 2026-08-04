*> vybe-test: cobol/category_pointers/test_pointers_allocate
*> origin: languages/cobol/tests/cobol/test_category_pointers.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PTR-ALLOC.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 WS-PTR USAGE POINTER.
       LINKAGE SECTION.
       01 LK-VAL PIC X(10).
       PROCEDURE DIVISION.
           ALLOCATE LK-VAL RETURNING WS-PTR.
           SET ADDRESS OF LK-VAL TO WS-PTR.
           MOVE "ALLOCATED" TO LK-VAL.
           DISPLAY LK-VAL.
    MOVE SPACES TO WS-VYBE-L
    STRING LK-VAL DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ALLOCATED "
        DISPLAY "FAIL: want [ALLOCATED ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           FREE WS-PTR.
           STOP RUN.

