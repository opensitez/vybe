*> vybe-test: cobol/qualified_names_of_clause/qualified_in_inspect
*> origin: languages/cobol/tests/cobol/test_qualified_names_of_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 STR-A.
   05 DATA PIC X(10) VALUE "HELLO".
01 STR-B.
   05 DATA PIC X(10) VALUE "WORLD".
01 CNT PIC 9(2) VALUE 0.
PROCEDURE DIVISION.
    INSPECT DATA OF STR-A TALLYING CNT FOR ALL "L".
    STOP RUN.

