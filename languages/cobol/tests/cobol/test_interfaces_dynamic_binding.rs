use super::helpers::compile_ok;

#[test]
fn interface_decl_a_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nINTERFACE-ID. IA.\nMETHOD-ID. M1.\nPROCEDURE DIVISION.\nEND METHOD M1.\nEND INTERFACE IA.",
    );
}
#[test]
fn interface_decl_b_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nINTERFACE-ID. IB.\nMETHOD-ID. M2.\nPROCEDURE DIVISION.\nEND METHOD M2.\nEND INTERFACE IB.",
    );
}
#[test]
fn class_impl_ia_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. CA IMPLEMENTS IA.\nOBJECT.\nMETHOD-ID. M1.\nPROCEDURE DIVISION.\n    DISPLAY \"1\".\nEND METHOD M1.\nEND OBJECT.\nEND CLASS CA.",
    );
}
#[test]
fn class_impl_ib_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. CB IMPLEMENTS IB.\nOBJECT.\nMETHOD-ID. M2.\nPROCEDURE DIVISION.\n    DISPLAY \"2\".\nEND METHOD M2.\nEND OBJECT.\nEND CLASS CB.",
    );
}
#[test]
fn class_inherit_and_impl_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. CC INHERITS FROM CA IMPLEMENTS IB.\nOBJECT.\nMETHOD-ID. M2.\nPROCEDURE DIVISION.\n    DISPLAY \"3\".\nEND METHOD M2.\nEND OBJECT.\nEND CLASS CC.",
    );
}
#[test]
fn dynamic_dispatch_call_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 O USAGE POINTER.\nPROCEDURE DIVISION.\n    INVOKE O M1.\n    STOP RUN.",
    );
}
#[test]
fn dynamic_dispatch_returning_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 O USAGE POINTER.\n01 R PIC 9.\nPROCEDURE DIVISION.\n    INVOKE O M1 RETURNING R.\n    STOP RUN.",
    );
}
#[test]
fn dynamic_set_method_ref_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 M PIC X(2) VALUE \"M1\".\nPROCEDURE DIVISION.\n    CALL \"BIND-METHOD\" USING M.\n    STOP RUN.",
    );
}
#[test]
fn dynamic_call_bound_ref_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"CALL-BOUND\".\n    STOP RUN.",
    );
}
#[test]
fn dynamic_interface_check_call_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 O USAGE POINTER.\n01 R PIC 9.\nPROCEDURE DIVISION.\n    CALL \"IMPLEMENTS\" USING O R.\n    STOP RUN.",
    );
}
#[test]
fn dynamic_cast_call_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 O USAGE POINTER.\nPROCEDURE DIVISION.\n    CALL \"DYN-CAST\" USING O.\n    STOP RUN.",
    );
}
#[test]
fn dynamic_proxy_call_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"MAKE-PROXY\".\n    STOP RUN.",
    );
}
#[test]
fn interface_method_map_call_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"MAP-IFACE\".\n    STOP RUN.",
    );
}
#[test]
fn dynamic_invoke_by_name_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 O USAGE POINTER.\n01 N PIC X(10) VALUE \"M1\".\nPROCEDURE DIVISION.\n    CALL \"INVOKE-NAME\" USING O N.\n    STOP RUN.",
    );
}
#[test]
fn dynamic_method_cache_call_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"METHOD-CACHE\".\n    STOP RUN.",
    );
}
#[test]
fn dynamic_virtual_dispatch_call_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"VIRTUAL-DISPATCH\".\n    STOP RUN.",
    );
}
#[test]
fn dynamic_trait_like_call_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"TRAIT-CALL\".\n    STOP RUN.",
    );
}
#[test]
fn dynamic_reflection_call_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"REFLECT\".\n    STOP RUN.",
    );
}

#[test]
fn interface_with_multiple_methods_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nINTERFACE-ID. IMULTI.\nMETHOD-ID. M1.\nPROCEDURE DIVISION.\nEND METHOD M1.\nMETHOD-ID. M2.\nPROCEDURE DIVISION.\nEND METHOD M2.\nEND INTERFACE IMULTI.",
    );
}

#[test]
fn class_implements_with_override_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. BASE.\nOBJECT.\nMETHOD-ID. DO.\nPROCEDURE DIVISION.\n    DISPLAY \"BASE\".\nEND METHOD DO.\nEND OBJECT.\nEND CLASS BASE.\n\nIDENTIFICATION DIVISION.\nCLASS-ID. DERIVED INHERITS FROM BASE IMPLEMENTS IMULTI.\nOBJECT.\nMETHOD-ID. DO.\nPROCEDURE DIVISION.\n    DISPLAY \"DERIVED\".\nEND METHOD DO.\nMETHOD-ID. M1.\nPROCEDURE DIVISION.\n    DISPLAY \"M1\".\nEND METHOD M1.\nMETHOD-ID. M2.\nPROCEDURE DIVISION.\n    DISPLAY \"M2\".\nEND METHOD M2.\nEND OBJECT.\nEND CLASS DERIVED.",
    );
}

#[test]
fn interface_impl_invoke_chain_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. CALLER.\nOBJECT.\nMETHOD-ID. HANDLE.\nPROCEDURE DIVISION USING WS-MSG.\n    DISPLAY WS-MSG.\nEND METHOD HANDLE.\nEND OBJECT.\nEND CLASS CALLER.",
    );
}
