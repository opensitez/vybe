*> vybe-test: cobol/oo_properties_and_attributes/class_read_only_property_compiles
*> origin: languages/cobol/tests/cobol/test_oo_properties_and_attributes.rs

IDENTIFICATION DIVISION.
CLASS-ID. READ-ONLY.
OBJECT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SKU PIC X(20) VALUE "SKU-001".
METHOD-ID. GET-SKU PROPERTY GET.
PROCEDURE DIVISION RETURNING WS-RESULT.
    MOVE WS-SKU TO WS-RESULT.
END METHOD GET-SKU.
END OBJECT.
END CLASS READ-ONLY.

