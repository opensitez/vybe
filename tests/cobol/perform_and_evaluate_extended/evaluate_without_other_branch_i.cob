*> vybe-test: cobol/perform_and_evaluate_extended/evaluate_without_other_branch_is_accepted
*> origin: languages/cobol/tests/cobol/test_perform_and_evaluate_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-CODE PIC 9 VALUE 1.
PROCEDURE DIVISION.

    EVALUATE WS-CODE
        WHEN 1
            DISPLAY "ONE"
        WHEN 2
            DISPLAY "TWO"
    END-EVALUATE.
    STOP RUN.

