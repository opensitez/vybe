*> vybe-test: cobol/special_registers_detail/address_of_in_loop_compiles
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 ITEM PIC X(10) VALUE "DATA".
01 P USAGE POINTER.
01 I PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM UNTIL I >= 3
        ADD 1 TO I
        SET P TO ADDRESS OF ITEM
    END-PERFORM.
    STOP RUN.

