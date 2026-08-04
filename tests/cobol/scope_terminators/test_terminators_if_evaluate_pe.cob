*> vybe-test: cobol/scope_terminators/test_terminators_if_evaluate_perform
*> origin: languages/cobol/tests/cobol/test_scope_terminators.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9 VALUE 5.
01 WS-I PIC 9.
PROCEDURE DIVISION.

    IF WS-A > 0
        EVALUATE WS-A
            WHEN 5
                PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 3
                    DISPLAY WS-I
                END-PERFORM
        END-EVALUATE
    END-IF.
    STOP RUN.

