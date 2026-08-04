*> vybe-test: cobol/interfaces_dynamic_binding/interface_impl_invoke_chain_compiles
*> origin: languages/cobol/tests/cobol/test_interfaces_dynamic_binding.rs
IDENTIFICATION DIVISION.
CLASS-ID. CALLER.
OBJECT.
METHOD-ID. HANDLE.
PROCEDURE DIVISION USING WS-MSG.
    DISPLAY WS-MSG.
END METHOD HANDLE.
END OBJECT.
END CLASS CALLER.

