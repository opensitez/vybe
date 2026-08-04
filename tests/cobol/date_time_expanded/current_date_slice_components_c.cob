*> vybe-test: cobol/date_time_expanded/current_date_slice_components_compiles
*> origin: languages/cobol/tests/cobol/test_date_time_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 CD PIC X(21).
01 Y PIC X(4).
PROCEDURE DIVISION.
    MOVE FUNCTION CURRENT-DATE TO CD.
    MOVE CD(1:4) TO Y.
    STOP RUN.

