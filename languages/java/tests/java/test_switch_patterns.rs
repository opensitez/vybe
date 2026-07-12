use crate::helpers::{run_in_main, run_main};

#[test]
fn switch_type_pattern_string_binds_and_prints_value() {
    let out = run_main(
        "Object o = \"hi\"; switch (o) { case String s -> System.out.println(s); default -> System.out.println(\"other\"); }",
    );
    assert_eq!(out, vec!["hi"]);
}

#[test]
fn switch_type_pattern_integer_binds_numeric_value() {
    let out = run_main(
        "Object o = 42; switch (o) { case Integer i -> System.out.println(i); default -> System.out.println(\"other\"); }",
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn switch_type_pattern_default_when_no_type_matches() {
    let out = run_main(
        "Object o = 3.14; switch (o) { case String s -> System.out.println(\"str\"); case Integer i -> System.out.println(\"int\"); default -> System.out.println(\"other\"); }",
    );
    assert_eq!(out, vec!["other"]);
}

#[test]
fn switch_null_case_matches_explicit_null_selector() {
    let out = run_main(
        "Object o = null; switch (o) { case null -> System.out.println(\"null\"); case String s -> System.out.println(\"str\"); default -> System.out.println(\"other\"); }",
    );
    assert_eq!(out, vec!["null"]);
}

#[test]
fn switch_null_case_before_string_pattern() {
    let out = run_main(
        "String s = null; switch (s) { case null -> System.out.println(\"nil\"); case String t -> System.out.println(t); }",
    );
    assert_eq!(out, vec!["nil"]);
}

#[test]
fn switch_guarded_pattern_string_length_positive() {
    let out = run_main(
        "Object o = \"alpha\"; switch (o) { case String s when s.length() > 0 -> System.out.println(\"nonempty\"); default -> System.out.println(\"empty\"); }",
    );
    assert_eq!(out, vec!["nonempty"]);
}

#[test]
fn switch_guarded_pattern_string_length_zero_falls_to_default() {
    let out = run_main(
        "Object o = \"\"; switch (o) { case String s when s.length() > 0 -> System.out.println(\"nonempty\"); default -> System.out.println(\"empty\"); }",
    );
    assert_eq!(out, vec!["empty"]);
}

#[test]
fn switch_guarded_pattern_prefix_match() {
    let out = run_main(
        "Object o = \"java21\"; switch (o) { case String s when s.startsWith(\"java\") -> System.out.println(\"j\"); default -> System.out.println(\"x\"); }",
    );
    assert_eq!(out, vec!["j"]);
}

#[test]
fn switch_guarded_pattern_numeric_comparison_on_integer() {
    let out = run_main(
        "Object o = 9; switch (o) { case Integer i when i > 5 -> System.out.println(\"big\"); case Integer i -> System.out.println(\"small\"); default -> System.out.println(\"other\"); }",
    );
    assert_eq!(out, vec!["big"]);
}

#[test]
fn switch_guarded_pattern_second_arm_when_first_guard_fails() {
    let out = run_main(
        "Object o = 2; switch (o) { case Integer i when i > 5 -> System.out.println(\"big\"); case Integer i -> System.out.println(\"small\"); default -> System.out.println(\"other\"); }",
    );
    assert_eq!(out, vec!["small"]);
}

#[test]
fn switch_expression_type_pattern_yields_bound_string() {
    let out = run_main(
        "Object o = \"ok\"; String label = switch (o) { case String s -> s.toUpperCase(); default -> \"NA\"; }; System.out.println(label);",
    );
    assert_eq!(out, vec!["OK"]);
}

#[test]
fn switch_expression_guarded_pattern_selects_matching_arm() {
    let out = run_main(
        "Object o = 7; int code = switch (o) { case Integer i when i % 2 == 0 -> 0; case Integer i -> 1; default -> -1; }; System.out.println(code);",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn switch_type_pattern_long_binding() {
    let out = run_main(
        "Object o = 100L; switch (o) { case Long l -> System.out.println(l); default -> System.out.println(\"other\"); }",
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn switch_type_pattern_double_binding() {
    let out = run_main(
        "Object o = 2.5; switch (o) { case Double d -> System.out.println(d); default -> System.out.println(\"other\"); }",
    );
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn switch_type_pattern_boolean_binding() {
    let out = run_main(
        "Object o = true; switch (o) { case Boolean b -> System.out.println(b); default -> System.out.println(\"other\"); }",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn switch_type_pattern_character_binding() {
    let out = run_main(
        "Object o = 'Z'; switch (o) { case Character c -> System.out.println(c); default -> System.out.println(\"other\"); }",
    );
    assert_eq!(out, vec!["Z"]);
}

#[test]
fn switch_type_pattern_uses_bound_variable_in_body() {
    let out = run_main(
        "Object o = \"ab\"; switch (o) { case String s -> System.out.println(s.length()); default -> System.out.println(0); }",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn switch_type_pattern_string_colon_style_with_break() {
    let out = run_main(
        "Object o = \"go\"; switch (o) { case String s: System.out.println(s); break; default: System.out.println(\"x\"); }",
    );
    assert_eq!(out, vec!["go"]);
}

#[test]
fn switch_guarded_pattern_colon_style_with_break() {
    let out = run_main(
        "Object o = \"zip\"; switch (o) { case String s when s.length() == 3: System.out.println(\"tri\"); break; default: System.out.println(\"no\"); }",
    );
    assert_eq!(out, vec!["tri"]);
}

#[test]
fn switch_type_pattern_subclass_via_object_reference() {
    let types = r#"
        static class Box { String id() { return "box"; } }
    "#;
    let out = run_in_main(
        "Object o = new Box(); switch (o) { case Box b -> System.out.println(b.id()); default -> System.out.println(\"other\"); }",
        types,
    );
    assert_eq!(out, vec!["box"]);
}

#[test]
fn switch_type_pattern_interface_implementor() {
    let types = r#"
        interface Named { String name(); }
        static class Item implements Named { public String name() { return "item"; } }
    "#;
    let out = run_in_main(
        "Object o = new Item(); switch (o) { case Named n -> System.out.println(n.name()); default -> System.out.println(\"x\"); }",
        types,
    );
    assert_eq!(out, vec!["item"]);
}

#[test]
fn switch_null_case_in_expression_returns_sentinel() {
    let out = run_main(
        "Object o = null; int n = switch (o) { case null -> -1; case Integer i -> i; default -> 0; }; System.out.println(n);",
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn switch_multiple_type_patterns_first_match_wins() {
    let out = run_main(
        "Object o = \"5\"; switch (o) { case String s -> System.out.println(\"s\"); case Integer i -> System.out.println(\"i\"); default -> System.out.println(\"d\"); }",
    );
    assert_eq!(out, vec!["s"]);
}

#[test]
fn switch_guarded_pattern_equals_literal() {
    let out = run_main(
        "Object o = \"yes\"; switch (o) { case String s when s.equals(\"yes\") -> System.out.println(1); default -> System.out.println(0); }",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn switch_guarded_pattern_negated_condition() {
    let out = run_main(
        "Object o = \"no\"; switch (o) { case String s when !s.equals(\"yes\") -> System.out.println(1); default -> System.out.println(0); }",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn switch_type_pattern_nested_switch_on_bound_value() {
    let out = run_main(
        "Object o = \"ab\"; switch (o) { case String s -> { switch (s.length()) { case 2 -> System.out.println(\"two\"); default -> System.out.println(\"other\"); } } default -> System.out.println(\"x\"); }",
    );
    assert_eq!(out, vec!["two"]);
}

#[test]
fn switch_type_pattern_array_int_binding() {
    let out = run_main(
        "Object o = new int[] {1, 2, 3}; switch (o) { case int[] a -> System.out.println(a.length); default -> System.out.println(0); }",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn switch_type_pattern_string_array_binding() {
    let out = run_main(
        "Object o = new String[] {\"a\", \"b\"}; switch (o) { case String[] a -> System.out.println(a.length); default -> System.out.println(0); }",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn switch_guarded_pattern_chained_length_and_starts_with() {
    let out = run_main(
        "Object o = \"java\"; switch (o) { case String s when s.length() == 4 && s.startsWith(\"j\") -> System.out.println(\"ok\"); default -> System.out.println(\"no\"); }",
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn switch_type_pattern_enum_constant_via_object() {
    let types = r#"
        enum Level { LOW, HIGH }
    "#;
    let out = run_in_main(
        "Object o = Level.HIGH; switch (o) { case Level l -> System.out.println(l == Level.HIGH); default -> System.out.println(false); }",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn switch_guarded_pattern_integer_even_check() {
    let out = run_main(
        "Object o = 8; switch (o) { case Integer i when i % 2 == 0 -> System.out.println(\"even\"); case Integer i -> System.out.println(\"odd\"); default -> System.out.println(\"x\"); }",
    );
    assert_eq!(out, vec!["even"]);
}

#[test]
fn switch_guarded_pattern_integer_odd_check() {
    let out = run_main(
        "Object o = 7; switch (o) { case Integer i when i % 2 == 0 -> System.out.println(\"even\"); case Integer i -> System.out.println(\"odd\"); default -> System.out.println(\"x\"); }",
    );
    assert_eq!(out, vec!["odd"]);
}

#[test]
fn switch_type_pattern_record_destructure_via_binding() {
    let types = r#"
        record Point(int x, int y) {}
    "#;
    let out = run_in_main(
        "Object o = new Point(3, 4); switch (o) { case Point p -> System.out.println(p.x() + p.y()); default -> System.out.println(0); }",
        types,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn switch_null_case_only_arm_on_nullable_string() {
    let out = run_main(
        "String s = null; String out = switch (s) { case null -> \"nil\"; case String t -> t; }; System.out.println(out);",
    );
    assert_eq!(out, vec!["nil"]);
}

#[test]
fn switch_type_pattern_default_after_null_and_string() {
    let out = run_main(
        "Object o = 1.0f; switch (o) { case null -> System.out.println(\"n\"); case String s -> System.out.println(\"s\"); default -> System.out.println(\"d\"); }",
    );
    assert_eq!(out, vec!["d"]);
}

#[test]
fn switch_guarded_pattern_case_insensitive_prefix() {
    let out = run_main(
        "Object o = \"Java\"; switch (o) { case String s when s.toLowerCase().startsWith(\"java\") -> System.out.println(1); default -> System.out.println(0); }",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn switch_type_pattern_block_body_prints_twice() {
    let out = run_main(
        "Object o = \"x\"; switch (o) { case String s -> { System.out.println(s); System.out.println(s.length()); } default -> System.out.println(\"z\"); }",
    );
    assert_eq!(out, vec!["x", "1"]);
}

#[test]
fn switch_expression_guarded_string_not_blank() {
    let out = run_main(
        "Object o = \"  a  \"; String tag = switch (o) { case String s when !s.isBlank() -> \"text\"; default -> \"blank\"; }; System.out.println(tag);",
    );
    assert_eq!(out, vec!["text"]);
}

#[test]
fn switch_type_pattern_short_wrapper_binding() {
    let out = run_main(
        "Object o = (short) 5; switch (o) { case Short s -> System.out.println(s); default -> System.out.println(0); }",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn switch_type_pattern_byte_wrapper_binding() {
    let out = run_main(
        "Object o = (byte) 9; switch (o) { case Byte b -> System.out.println(b); default -> System.out.println(0); }",
    );
    assert_eq!(out, vec!["9"]);
}
