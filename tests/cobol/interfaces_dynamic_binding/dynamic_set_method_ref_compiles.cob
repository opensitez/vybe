*> vybe-test: cobol/interfaces_dynamic_binding/dynamic_set_method_ref_compiles
*> origin: languages/cobol/tests/cobol/test_interfaces_dynamic_binding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 M PIC X(2) VALUE "M1".
PROCEDURE DIVISION.
    CALL "BIND-METHOD" USING M.
    STOP RUN.

