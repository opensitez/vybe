*> vybe-test: cobol/properties_attributes_accessors/prop_name_property_compiles
*> origin: languages/cobol/tests/cobol/test_properties_attributes_accessors.rs
IDENTIFICATION DIVISION.
CLASS-ID. P4.
OBJECT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC X(20).
METHOD-ID. GET-N PROPERTY GET.
PROCEDURE DIVISION RETURNING R.
    MOVE N TO R.
END METHOD GET-N.
END OBJECT.
END CLASS P4.

