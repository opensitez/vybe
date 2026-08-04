*> vybe-test: cobol/basic_classes_oop/invoke_method_on_object_reference_compiles
*> origin: languages/cobol/tests/cobol/test_basic_classes_oop.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. INVOKE-CLASS.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-OBJ USAGE POINTER.
01 WS-OUT PIC 9(4).
PROCEDURE DIVISION.
    INVOKE WS-OBJ GET-CODE RETURNING WS-OUT.
    DISPLAY WS-OUT.
    STOP RUN.

