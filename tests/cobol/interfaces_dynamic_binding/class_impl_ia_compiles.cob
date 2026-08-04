*> vybe-test: cobol/interfaces_dynamic_binding/class_impl_ia_compiles
*> origin: languages/cobol/tests/cobol/test_interfaces_dynamic_binding.rs
IDENTIFICATION DIVISION.
CLASS-ID. CA IMPLEMENTS IA.
OBJECT.
METHOD-ID. M1.
PROCEDURE DIVISION.
    DISPLAY "1".
END METHOD M1.
END OBJECT.
END CLASS CA.

