*> vybe-test: cobol/properties_attributes_accessors/prop_class_with_compute_compiles
*> origin: languages/cobol/tests/cobol/test_properties_attributes_accessors.rs
IDENTIFICATION DIVISION.
CLASS-ID. P12.
OBJECT.
METHOD-ID. VAL.
PROCEDURE DIVISION RETURNING R.
    COMPUTE R = 2 + 3.
END METHOD VAL.
END OBJECT.
END CLASS P12.

