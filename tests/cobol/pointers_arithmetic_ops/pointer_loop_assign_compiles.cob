*> vybe-test: cobol/pointers_arithmetic_ops/pointer_loop_assign_compiles
*> origin: languages/cobol/tests/cobol/test_pointers_arithmetic_ops.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 P USAGE POINTER.
01 I PIC 9 VALUE 1.
PROCEDURE DIVISION.
    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3
        SET P TO NULL
    END-PERFORM.
    STOP RUN.

