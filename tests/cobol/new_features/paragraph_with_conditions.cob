*> vybe-test: cobol/new_features/paragraph_with_conditions
*> origin: languages/cobol/tests/cobol/test_new_features.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. PARACONDN.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TYPE   PIC X(1) VALUE "A".
   88 IS-TYPE-A VALUE "A".
   88 IS-TYPE-B VALUE "B".
   88 IS-TYPE-C VALUE "C".
PROCEDURE DIVISION.
    PERFORM PROCESS-PARA.
    STOP RUN.
PROCESS-PARA.
    IF IS-TYPE-A
        DISPLAY "Processing type A"
    ELSE
        IF IS-TYPE-B
            DISPLAY "Processing type B"
        ELSE
            DISPLAY "Processing other type"
        END-IF
    END-IF.

