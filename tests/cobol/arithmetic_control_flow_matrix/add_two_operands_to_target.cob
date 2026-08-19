*> vybe-test: cobol/arithmetic_control_flow_matrix/add_two_operands_to_target
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 2.
01 B PIC 9(3) VALUE 3.
01 R PIC 9(3) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    ADD A B TO R.
    DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "005"
        DISPLAY "FAIL: want [005] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

