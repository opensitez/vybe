*> vybe-test: cobol/qualified_names_of_clause/qualified_nested_four_levels
*> origin: languages/cobol/tests/cobol/test_qualified_names_of_clause.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 L1.
   05 L2.
      10 L3.
         15 NAME PIC X(5) VALUE "COBOL".
PROCEDURE DIVISION.
    DISPLAY NAME OF L3 OF L2 OF L1.
    STOP RUN.

