*> vybe-test: cobol/scope_terminators_nesting/scope_end_search_compiles
*> origin: languages/cobol/tests/cobol/test_scope_terminators_nesting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC X OCCURS 5 TIMES.
PROCEDURE DIVISION.
    SEARCH E
        AT END
            DISPLAY "NOT FOUND"
        WHEN E(1) = "A"
            DISPLAY "FOUND"
    END-SEARCH.
    STOP RUN.

