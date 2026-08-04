*> vybe-test: cobol/pointers/pointer_move_between_pointer_variables_compiles
*> origin: languages/cobol/tests/cobol/test_pointers.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P1 USAGE POINTER.
01 P2 USAGE POINTER.
01 BUF PIC X(4) VALUE "DATA".
PROCEDURE DIVISION.
    SET P1 TO ADDRESS OF BUF.
    MOVE P1 TO P2.
    STOP RUN.

