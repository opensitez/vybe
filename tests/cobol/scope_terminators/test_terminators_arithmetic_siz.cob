*> vybe-test: cobol/scope_terminators/test_terminators_arithmetic_size_error
*> origin: languages/cobol/tests/cobol/test_scope_terminators.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9 VALUE 9.
PROCEDURE DIVISION.

    ADD 1 TO WS-A
        ON SIZE ERROR
            DISPLAY "OVERFLOW"
    END-ADD.
    
    SUBTRACT 1 FROM WS-A
        ON SIZE ERROR
            DISPLAY "OVERFLOW"
    END-SUBTRACT.
    
    MULTIPLY 2 BY WS-A
        ON SIZE ERROR
            DISPLAY "OVERFLOW"
    END-MULTIPLY.
    
    DIVIDE 2 INTO WS-A
        ON SIZE ERROR
            DISPLAY "OVERFLOW"
    END-DIVIDE.
    
    COMPUTE WS-A = WS-A * 2
        ON SIZE ERROR
            DISPLAY "OVERFLOW"
    END-COMPUTE.
    STOP RUN.

