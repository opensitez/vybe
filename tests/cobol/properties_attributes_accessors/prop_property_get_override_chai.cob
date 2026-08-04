*> vybe-test: cobol/properties_attributes_accessors/prop_property_get_override_chain_compiles
*> origin: languages/cobol/tests/cobol/test_properties_attributes_accessors.rs
IDENTIFICATION DIVISION.
CLASS-ID. P15 INHERITS FROM P14.
OBJECT.
METHOD-ID. GET-Z PROPERTY GET.
PROCEDURE DIVISION RETURNING R.
    MOVE 1 TO R.
END METHOD GET-Z.
END OBJECT.
END CLASS P15.

