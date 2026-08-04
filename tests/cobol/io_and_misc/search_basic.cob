*> vybe-test: cobol/io_and_misc/search_basic
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TBL.
   05 ITEM PIC X(10) OCCURS 10 TIMES.
PROCEDURE DIVISION.
    SEARCH ITEM
        AT END
            DISPLAY "Not found"
        WHEN ITEM(1) = "A"
            DISPLAY "Found"
    END-SEARCH.
    STOP RUN.

