*> vybe-test: cobol/enum_like_states/condition_name_false_assignment_compiles
*> origin: languages/cobol/tests/cobol/test_enum_like_states.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-FLAG PIC 9 VALUE 1.
   88 IS-ON VALUE 1.
   88 IS-OFF VALUE 0.
PROCEDURE DIVISION.
    SET IS-OFF TO TRUE.
    IF IS-OFF DISPLAY "OFF" END-IF.
    STOP RUN.

