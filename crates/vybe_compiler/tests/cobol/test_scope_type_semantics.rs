use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test] fn type_alpha_picx_compiles() { compile_ok(&p("01 A PIC X(10).", "    MOVE \"A\" TO A.")); }
#[test] fn type_numeric_pic9_compiles() { compile_ok(&p("01 A PIC 9(5).", "    MOVE 12 TO A.")); }
#[test] fn type_signed_numeric_compiles() { compile_ok(&p("01 A PIC S9(5).", "    MOVE -12 TO A.")); }
#[test] fn type_decimal_v_compiles() { compile_ok(&p("01 A PIC 9(3)V99.", "    MOVE 123.45 TO A.")); }
#[test] fn type_binary_usage_compiles() { compile_ok(&p("01 A PIC 9(5) USAGE BINARY.", "    ADD 1 TO A.")); }
#[test] fn type_comp_usage_compiles() { compile_ok(&p("01 A PIC 9(5) USAGE COMP.", "    ADD 1 TO A.")); }
#[test] fn type_comp3_usage_compiles() { compile_ok(&p("01 A PIC 9(5) USAGE COMP-3.", "    ADD 1 TO A.")); }
#[test] fn type_pointer_usage_compiles() { compile_ok(&p("01 P USAGE POINTER.", "    SET P TO NULL.")); }
#[test] fn type_function_pointer_usage_compiles() { compile_ok(&p("01 P USAGE FUNCTION-POINTER.", "    DISPLAY \"P\".")); }
#[test] fn type_procedure_pointer_usage_compiles() { compile_ok(&p("01 P USAGE PROCEDURE-POINTER.", "    DISPLAY \"P\".")); }
#[test] fn scope_nested_if_compiles() { compile_ok(&p("01 X PIC 9 VALUE 1.", "    IF X = 1\n        IF X > 0 DISPLAY \"Y\" END-IF\n    END-IF.")); }
#[test] fn scope_evaluate_inside_if_compiles() { compile_ok(&p("01 X PIC 9 VALUE 2.", "    IF X > 0\n        EVALUATE X\n            WHEN 1 DISPLAY \"A\"\n            WHEN OTHER DISPLAY \"B\"\n        END-EVALUATE\n    END-IF.")); }
#[test] fn scope_perform_inside_if_compiles() { compile_ok(&p("01 X PIC 9 VALUE 1.", "    IF X = 1\n        PERFORM 2 TIMES DISPLAY \"L\" END-PERFORM\n    END-IF.")); }
#[test] fn scope_if_inside_perform_compiles() { compile_ok(&p("01 X PIC 9 VALUE 1.", "    PERFORM 2 TIMES\n        IF X = 1 DISPLAY \"A\" END-IF\n    END-PERFORM.")); }
#[test] fn scope_group_data_compiles() { compile_ok(&p("01 G.\n   05 A PIC X(3).\n   05 B PIC 9(2).", "    MOVE \"ABC\" TO A.")); }
#[test] fn scope_redefines_data_compiles() { compile_ok(&p("01 B PIC X(10).\n01 N REDEFINES B PIC 9(10).", "    MOVE 1 TO N.")); }
#[test] fn scope_occurs_data_compiles() { compile_ok(&p("01 T PIC X(2) OCCURS 3 TIMES.", "    MOVE \"AA\" TO T(1).")); }
#[test] fn scope_condition_name_compiles() { compile_ok(&p("01 F PIC 9.\n   88 ONN VALUE 1.", "    SET ONN TO TRUE.")); }
