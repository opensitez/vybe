*> vybe-test: cobol/condition_names_level88_states/condition_name_range_runtime_transitions
*> origin: languages/cobol/tests/cobol/test_condition_names_level88_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 ST PIC 9(2) VALUE 0.
   88 LOW VALUE 1 THRU 5.
   88 MID VALUE 6 THRU 10.
   88 HIGH VALUE 11 THRU 20.
PROCEDURE DIVISION.
    SET LOW TO TRUE
    IF LOW DISPLAY "LOW" ELSE DISPLAY "NOT-LOW" END-IF
    SET HIGH TO TRUE
    IF HIGH DISPLAY "HIGH" ELSE DISPLAY "NOT-HIGH" END-IF
    STOP RUN.

