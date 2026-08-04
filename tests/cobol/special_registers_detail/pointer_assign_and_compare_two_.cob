*> vybe-test: cobol/special_registers_detail/pointer_assign_and_compare_two_pointers
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(5) VALUE "HELLO".
01 P1 USAGE POINTER.
01 P2 USAGE POINTER.
PROCEDURE DIVISION.
    SET P1 TO ADDRESS OF A.
    SET P2 TO ADDRESS OF A.
    IF P1 = P2
        DISPLAY "SAME"
    END-IF.
    STOP RUN.

