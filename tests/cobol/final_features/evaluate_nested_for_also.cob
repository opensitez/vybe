*> vybe-test: cobol/final_features/evaluate_nested_for_also
*> origin: languages/cobol/tests/cobol/test_final_features.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. EVALALSO.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STATUS PIC X(1) VALUE "A".
01 WS-REGION PIC 9(1) VALUE 1.
PROCEDURE DIVISION.
    EVALUATE WS-STATUS
        WHEN "A"
            EVALUATE WS-REGION
                WHEN 1
                    DISPLAY "Active Region 1"
                WHEN 2
                    DISPLAY "Active Region 2"
            END-EVALUATE
        WHEN "I"
            DISPLAY "Inactive"
    END-EVALUATE.
    STOP RUN.

