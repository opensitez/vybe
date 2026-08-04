*> vybe-test: cobol/new_features/multi_condition_evaluate
*> origin: languages/cobol/tests/cobol/test_new_features.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. MULTIEVAL.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STATUS PIC X(1) VALUE "A".
01 WS-REGION PIC 9(1) VALUE 1.
PROCEDURE DIVISION.
    EVALUATE WS-STATUS
        WHEN "A"
            EVALUATE WS-REGION
                WHEN 1
                    DISPLAY "Active - Region 1"
                WHEN 2
                    DISPLAY "Active - Region 2"
                WHEN OTHER
                    DISPLAY "Active - Other Region"
            END-EVALUATE
        WHEN "I"
            DISPLAY "Inactive"
        WHEN OTHER
            DISPLAY "Unknown"
    END-EVALUATE.
    STOP RUN.

