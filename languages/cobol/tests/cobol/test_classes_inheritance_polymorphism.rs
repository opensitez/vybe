use super::helpers::compile_ok;

#[test]
fn class_base_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. BASE-C.\nOBJECT.\nMETHOD-ID. M1.\nPROCEDURE DIVISION.\n    DISPLAY \"M1\".\nEND METHOD M1.\nEND OBJECT.\nEND CLASS BASE-C.",
    );
}
#[test]
fn class_derived_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. DERIVED-C INHERITS FROM BASE-C.\nOBJECT.\nMETHOD-ID. M1 OVERRIDE.\nPROCEDURE DIVISION.\n    DISPLAY \"OV\".\nEND METHOD M1.\nEND OBJECT.\nEND CLASS DERIVED-C.",
    );
}
#[test]
fn class_interface_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nINTERFACE-ID. IPRINT.\nMETHOD-ID. PRINT-SELF.\nPROCEDURE DIVISION.\nEND METHOD PRINT-SELF.\nEND INTERFACE IPRINT.",
    );
}
#[test]
fn class_implements_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. DOC IMPLEMENTS IPRINT.\nOBJECT.\nMETHOD-ID. PRINT-SELF.\nPROCEDURE DIVISION.\n    DISPLAY \"DOC\".\nEND METHOD PRINT-SELF.\nEND OBJECT.\nEND CLASS DOC.",
    );
}
#[test]
fn class_factory_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. UTIL.\nFACTORY.\nMETHOD-ID. BUILD.\nPROCEDURE DIVISION RETURNING WS-OBJ.\n    DISPLAY \"B\".\nEND METHOD BUILD.\nEND FACTORY.\nEND CLASS UTIL.",
    );
}
#[test]
fn class_with_data_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. PERSON-C.\nOBJECT.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-N PIC X(20).\nMETHOD-ID. SET-N.\nPROCEDURE DIVISION USING WS-IN.\n    MOVE WS-IN TO WS-N.\nEND METHOD SET-N.\nEND OBJECT.\nEND CLASS PERSON-C.",
    );
}
#[test]
fn class_property_get_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. C1.\nOBJECT.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-V PIC 9(3).\nMETHOD-ID. GET-V PROPERTY GET.\nPROCEDURE DIVISION RETURNING WS-R.\n    MOVE WS-V TO WS-R.\nEND METHOD GET-V.\nEND OBJECT.\nEND CLASS C1.",
    );
}
#[test]
fn class_property_set_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. C2.\nOBJECT.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-V PIC 9(3).\nMETHOD-ID. SET-V PROPERTY SET.\nPROCEDURE DIVISION USING WS-I.\n    MOVE WS-I TO WS-V.\nEND METHOD SET-V.\nEND OBJECT.\nEND CLASS C2.",
    );
}
#[test]
fn class_method_returning_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. C3.\nOBJECT.\nMETHOD-ID. CODE.\nPROCEDURE DIVISION RETURNING WS-R.\n    MOVE 7 TO WS-R.\nEND METHOD CODE.\nEND OBJECT.\nEND CLASS C3.",
    );
}
#[test]
fn class_invoke_pattern_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. P.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 O USAGE POINTER.\n01 R PIC 9(3).\nPROCEDURE DIVISION.\n    INVOKE O CODE RETURNING R.\n    STOP RUN.",
    );
}
#[test]
fn class_multiple_methods_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. C4.\nOBJECT.\nMETHOD-ID. A.\nPROCEDURE DIVISION.\n    DISPLAY \"A\".\nEND METHOD A.\nMETHOD-ID. B.\nPROCEDURE DIVISION.\n    DISPLAY \"B\".\nEND METHOD B.\nEND OBJECT.\nEND CLASS C4.",
    );
}
#[test]
fn class_second_level_inherit_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. C5 INHERITS FROM DERIVED-C.\nOBJECT.\nMETHOD-ID. X.\nPROCEDURE DIVISION.\n    DISPLAY \"X\".\nEND METHOD X.\nEND OBJECT.\nEND CLASS C5.",
    );
}
#[test]
fn interface_two_methods_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nINTERFACE-ID. I2.\nMETHOD-ID. M1.\nPROCEDURE DIVISION.\nEND METHOD M1.\nMETHOD-ID. M2.\nPROCEDURE DIVISION.\nEND METHOD M2.\nEND INTERFACE I2.",
    );
}
#[test]
fn class_implements_two_methods_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. C6 IMPLEMENTS I2.\nOBJECT.\nMETHOD-ID. M1.\nPROCEDURE DIVISION.\n    DISPLAY \"1\".\nEND METHOD M1.\nMETHOD-ID. M2.\nPROCEDURE DIVISION.\n    DISPLAY \"2\".\nEND METHOD M2.\nEND OBJECT.\nEND CLASS C6.",
    );
}
#[test]
fn class_static_like_factory_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. C7.\nFACTORY.\nMETHOD-ID. NEW-OBJ.\nPROCEDURE DIVISION RETURNING WS-O.\n    DISPLAY \"N\".\nEND METHOD NEW-OBJ.\nEND FACTORY.\nEND CLASS C7.",
    );
}
#[test]
fn class_display_in_method_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. C8.\nOBJECT.\nMETHOD-ID. SHOW.\nPROCEDURE DIVISION.\n    DISPLAY \"SHOW\".\nEND METHOD SHOW.\nEND OBJECT.\nEND CLASS C8.",
    );
}
#[test]
fn class_data_and_compute_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. C9.\nOBJECT.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 A PIC 9 VALUE 2.\n01 B PIC 9 VALUE 3.\nMETHOD-ID. SUM.\nPROCEDURE DIVISION RETURNING R.\n    COMPUTE R = A + B.\nEND METHOD SUM.\nEND OBJECT.\nEND CLASS C9.",
    );
}
#[test]
fn class_override_with_display_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. C10 INHERITS FROM BASE-C.\nOBJECT.\nMETHOD-ID. M1 OVERRIDE.\nPROCEDURE DIVISION.\n    DISPLAY \"OV2\".\nEND METHOD M1.\nEND OBJECT.\nEND CLASS C10.",
    );
}

#[test]
fn class_factory_default_value_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nCLASS-ID. C11.\nFACTORY.\nMETHOD-ID. NEW.\nPROCEDURE DIVISION RETURNING WS-VAL.\n    MOVE \"OKAY\" TO WS-VAL.\nEND METHOD NEW.\nEND FACTORY.\nEND CLASS C11.",
    );
}
