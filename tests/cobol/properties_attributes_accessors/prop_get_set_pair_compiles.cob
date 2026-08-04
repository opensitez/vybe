*> vybe-test: cobol/properties_attributes_accessors/prop_get_set_pair_compiles
*> origin: languages/cobol/tests/cobol/test_properties_attributes_accessors.rs
IDENTIFICATION DIVISION.
CLASS-ID. P3.
OBJECT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9.
METHOD-ID. GET-A PROPERTY GET.
PROCEDURE DIVISION RETURNING R.
    MOVE A TO R.
END METHOD GET-A.
METHOD-ID. SET-A PROPERTY SET.
PROCEDURE DIVISION USING I.
    MOVE I TO A.
END METHOD SET-A.
END OBJECT.
END CLASS P3.

