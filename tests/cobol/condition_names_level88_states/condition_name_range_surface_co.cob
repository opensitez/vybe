*> vybe-test: cobol/condition_names_level88_states/condition_name_range_surface_compiles
*> origin: languages/cobol/tests/cobol/test_condition_names_level88_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 ST PIC 9(2) VALUE 10.
   88 LOW VALUE 1 THRU 5.
   88 MID VALUE 6 THRU 10.
   88 HIGH VALUE 11 THRU 20.
PROCEDURE DIVISION.
    SET MID TO TRUE.
    STOP RUN.

