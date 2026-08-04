*> vybe-test: cobol/cobol2023_file_io/read_at_end
*> origin: languages/cobol/tests/cobol/test_cobol2023_file_io.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC PIC X(80).
01 WS-EOF PIC 9(1) VALUE 0.
PROCEDURE DIVISION.
    READ WS-FILE INTO WS-REC
        AT END SET WS-EOF TO TRUE
        NOT AT END DISPLAY WS-REC
    END-READ.
    STOP RUN.

