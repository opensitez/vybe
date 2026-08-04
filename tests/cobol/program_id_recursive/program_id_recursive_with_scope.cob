*> vybe-test: cobol/program_id_recursive/program_id_recursive_with_scope_terminators
*> origin: languages/cobol/tests/cobol/test_program_id_recursive.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    ADD 1 TO N
    END-ADD.
    SUBTRACT 1 FROM N
    END-SUBTRACT.
    COMPUTE N = N * 2
    END-COMPUTE.
    STOP RUN.

