*> vybe-test: cobol/special_registers_detail/special_registers_all_used_in_one_program
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 RETURN-CODE PIC 9(4) VALUE 0.
01 S PIC X(10) VALUE "AABABAB".
01 N PIC 9(5) VALUE 0.
01 P USAGE POINTER.
PROCEDURE DIVISION.
    MOVE 0 TO RETURN-CODE.
    MOVE 0 TO TALLY.
    INSPECT S TALLYING TALLY FOR ALL "A".
    SET P TO ADDRESS OF N.
    IF RETURN-CODE = 0 AND TALLY > 0 AND P NOT = NULL
        DISPLAY "ALL VALID"
    END-IF.
    STOP RUN.

