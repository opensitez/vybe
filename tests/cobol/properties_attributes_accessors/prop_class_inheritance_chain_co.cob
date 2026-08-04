*> vybe-test: cobol/properties_attributes_accessors/prop_class_inheritance_chain_compiles
*> origin: languages/cobol/tests/cobol/test_properties_attributes_accessors.rs
IDENTIFICATION DIVISION.
CLASS-ID. P13 INHERITS FROM P12.
OBJECT.
METHOD-ID. EXTRA.
PROCEDURE DIVISION.
    DISPLAY "E".
END METHOD EXTRA.
END OBJECT.
END CLASS P13.

