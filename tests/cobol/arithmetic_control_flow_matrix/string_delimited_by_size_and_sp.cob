*> vybe-test: cobol/arithmetic_control_flow_matrix/string_delimited_by_size_and_space
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(4) VALUE "ONE".
01 B PIC X(4) VALUE "TWO".
01 R PIC X(20) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    STRING A DELIMITED BY SPACE
           B DELIMITED BY SPACE
           INTO R.
    DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ONETWO"
        DISPLAY "FAIL: want [ONETWO] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

