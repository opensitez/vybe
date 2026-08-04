*> vybe-test: cobol/category_pointers/test_pointers_chain_address_of
*> origin: languages/cobol/tests/cobol/test_category_pointers.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. PTR-CHAIN.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 WS-SRC PIC X(5) VALUE "HELLO".
       01 WS-DST PIC X(5).
       01 WS-PTR USAGE POINTER.
       PROCEDURE DIVISION.
           SET WS-PTR TO ADDRESS OF WS-SRC.
           SET ADDRESS OF WS-DST TO WS-PTR.
           IF WS-DST = WS-SRC
               DISPLAY "POINTER COPY".
    MOVE SPACES TO WS-VYBE-L
    STRING "POINTER COPY" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "POINTER COPY"
        DISPLAY "FAIL: want [POINTER COPY] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           END-IF.
           STOP RUN.

