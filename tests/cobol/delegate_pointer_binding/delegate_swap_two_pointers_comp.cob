*> vybe-test: cobol/delegate_pointer_binding/delegate_swap_two_pointers_compiles
*> origin: languages/cobol/tests/cobol/test_delegate_pointer_binding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P1 USAGE IS PROCEDURE-POINTER.
01 P2 USAGE IS PROCEDURE-POINTER.
PROCEDURE DIVISION.
    SET P1 TO ENTRY "A".
    SET P2 TO ENTRY "B".
    MOVE P1 TO P2.
    STOP RUN.

