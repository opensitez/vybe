*> vybe-test: cobol/conditions_extended/if_with_elseif_style_branch_compiles
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 2.
PROCEDURE DIVISION.
    IF WS-A = 1
        DISPLAY "ONE"
    ELSE
        IF WS-A = 2
            DISPLAY "TWO"
        END-IF
    END-IF.
    STOP RUN.

