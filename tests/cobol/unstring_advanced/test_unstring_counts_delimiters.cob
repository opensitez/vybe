*> vybe-test: cobol/unstring_advanced/test_unstring_counts_delimiters
*> origin: languages/cobol/tests/cobol/test_unstring_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SRC PIC X(10) VALUE "A,BB;CCC".
01 WS-F1 PIC X(3).
01 WS-F2 PIC X(3).
01 WS-DEL1 PIC X.
01 WS-CNT1 PIC 99.
PROCEDURE DIVISION.

    UNSTRING WS-SRC DELIMITED BY "," OR ";"
        INTO WS-F1 DELIMITER IN WS-DEL1 COUNT IN WS-CNT1
             WS-F2.
    STOP RUN.

