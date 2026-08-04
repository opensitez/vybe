*> vybe-test: cobol/programs/table_processing
*> origin: languages/cobol/tests/cobol/test_programs.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. TABPROC.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TABLE.
   05 WS-ITEM PIC X(10) OCCURS 5 TIMES.
01 WS-I PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    MOVE "Apple"  TO WS-ITEM(1).
    MOVE "Banana" TO WS-ITEM(2).
    MOVE "Cherry" TO WS-ITEM(3).
    MOVE "Date"   TO WS-ITEM(4).
    MOVE "Fig"    TO WS-ITEM(5).
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 5
        DISPLAY "Item " WS-I ": " WS-ITEM(WS-I)
    END-PERFORM.
    STOP RUN.

