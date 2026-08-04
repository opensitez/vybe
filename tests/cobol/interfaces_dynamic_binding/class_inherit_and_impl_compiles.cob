*> vybe-test: cobol/interfaces_dynamic_binding/class_inherit_and_impl_compiles
*> origin: languages/cobol/tests/cobol/test_interfaces_dynamic_binding.rs
IDENTIFICATION DIVISION.
CLASS-ID. CC INHERITS FROM CA IMPLEMENTS IB.
OBJECT.
METHOD-ID. M2.
PROCEDURE DIVISION.
    DISPLAY "3".
END METHOD M2.
END OBJECT.
END CLASS CC.

