*> vybe-test: cobol/properties_attributes_accessors/prop_invoke_getter_pattern_compiles
*> origin: languages/cobol/tests/cobol/test_properties_attributes_accessors.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 O USAGE POINTER.
01 R PIC 9.
PROCEDURE DIVISION.
    INVOKE O GET-A RETURNING R.
    STOP RUN.

