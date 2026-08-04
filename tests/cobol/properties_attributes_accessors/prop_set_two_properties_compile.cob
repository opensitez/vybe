*> vybe-test: cobol/properties_attributes_accessors/prop_set_two_properties_compiles
*> origin: languages/cobol/tests/cobol/test_properties_attributes_accessors.rs
IDENTIFICATION DIVISION.
CLASS-ID. P7.
OBJECT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9.
01 B PIC 9.
METHOD-ID. SET-A PROPERTY SET.
PROCEDURE DIVISION USING I.
    MOVE I TO A.
END METHOD SET-A.
METHOD-ID. SET-B PROPERTY SET.
PROCEDURE DIVISION USING I.
    MOVE I TO B.
END METHOD SET-B.
END OBJECT.
END CLASS P7.

