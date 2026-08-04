*> vybe-test: cobol/interfaces_dynamic_binding/dynamic_cast_call_compiles
*> origin: languages/cobol/tests/cobol/test_interfaces_dynamic_binding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 O USAGE POINTER.
PROCEDURE DIVISION.
    CALL "DYN-CAST" USING O.
    STOP RUN.

