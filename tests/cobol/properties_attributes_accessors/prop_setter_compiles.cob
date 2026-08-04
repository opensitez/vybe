*> vybe-test: cobol/properties_attributes_accessors/prop_setter_compiles
*> origin: languages/cobol/tests/cobol/test_properties_attributes_accessors.rs
IDENTIFICATION DIVISION.
CLASS-ID. P2.
OBJECT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9.
METHOD-ID. SET-A PROPERTY SET.
PROCEDURE DIVISION USING I.
    MOVE I TO A.
END METHOD SET-A.
END OBJECT.
END CLASS P2.

