use crate::helpers::{run_in_main, run_main};

#[test]
fn instanceof_string_pattern_binds_and_calls_length() {
    let out = run_main(
        "Object o = \"abc\"; if (o instanceof String s) { System.out.println(s.length()); }",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn instanceof_string_pattern_uppercases_bound_variable() {
    let out = run_main(
        "Object o = \"java\"; if (o instanceof String s) { System.out.println(s.toUpperCase()); }",
    );
    assert_eq!(out, vec!["JAVA"]);
}

#[test]
fn instanceof_integer_pattern_binds_numeric_value() {
    let out = run_main("Object o = 15; if (o instanceof Integer n) { System.out.println(n + 1); }");
    assert_eq!(out, vec!["16"]);
}

#[test]
fn instanceof_negation_false_for_matching_string() {
    let out = run_main("Object o = \"x\"; System.out.println(!(o instanceof String));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn instanceof_negation_true_for_non_matching_type() {
    let out = run_main("Object o = 1; System.out.println(!(o instanceof String));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn instanceof_compound_and_length_guard() {
    let out = run_main(
        "Object o = \"hello\"; if (o instanceof String s && s.length() > 3) { System.out.println(\"long\"); } else { System.out.println(\"short\"); }",
    );
    assert_eq!(out, vec!["long"]);
}

#[test]
fn instanceof_compound_and_fails_when_length_too_small() {
    let out = run_main(
        "Object o = \"hi\"; if (o instanceof String s && s.length() > 3) { System.out.println(\"long\"); } else { System.out.println(\"short\"); }",
    );
    assert_eq!(out, vec!["short"]);
}

#[test]
fn instanceof_compound_or_accepts_integer_operand() {
    let out = run_main(
        "Object o = 4; if (o instanceof String s || o instanceof Integer n) { System.out.println(\"hit\"); } else { System.out.println(\"miss\"); }",
    );
    assert_eq!(out, vec!["hit"]);
}

#[test]
fn instanceof_compound_or_string_branch_prints_value() {
    let out = run_main(
        "Object o = \"ok\"; if (o instanceof String s || o instanceof Integer n) { System.out.println(o instanceof String ? s : \"num\"); } else { System.out.println(\"miss\"); }",
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn instanceof_pattern_false_skips_binding_block() {
    let out = run_main(
        "Object o = 9; if (o instanceof String s) { System.out.println(s); } else { System.out.println(\"skip\"); }",
    );
    assert_eq!(out, vec!["skip"]);
}

#[test]
fn instanceof_null_reference_is_always_false() {
    let out = run_main("Object o = null; System.out.println(o instanceof String);");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn instanceof_subclass_pattern_on_upcast_reference() {
    let types = r#"
        static class Animal {}
        static class Dog extends Animal { String bark() { return "woof"; } }
    "#;
    let out = run_in_main(
        "Animal a = new Dog(); if (a instanceof Dog d) { System.out.println(d.bark()); }",
        types,
    );
    assert_eq!(out, vec!["woof"]);
}

#[test]
fn instanceof_parent_true_for_child_instance() {
    let types = r#"
        static class Parent {}
        static class Child extends Parent {}
    "#;
    let out = run_in_main(
        "Child c = new Child(); System.out.println(c instanceof Parent);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn instanceof_sibling_type_is_false() {
    let types = r#"
        static class Alpha {}
        static class Beta {}
    "#;
    let out = run_in_main(
        "Alpha a = new Alpha(); System.out.println(a instanceof Beta);",
        types,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn instanceof_interface_on_concrete_implementor() {
    let types = r#"
        interface Printable { String render(); }
        static class Doc implements Printable { public String render() { return "doc"; } }
    "#;
    let out = run_in_main(
        "Printable p = new Doc(); if (p instanceof Doc d) { System.out.println(d.render()); }",
        types,
    );
    assert_eq!(out, vec!["doc"]);
}

#[test]
fn instanceof_array_pattern_on_int_array() {
    let out = run_main(
        "Object o = new int[] {1, 2}; if (o instanceof int[] a) { System.out.println(a.length); }",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn instanceof_array_pattern_on_string_array() {
    let out =
        run_main("Object o = new String[] {\"a\"}; System.out.println(o instanceof String[]);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn instanceof_pattern_in_ternary_selects_bound_branch() {
    let out = run_main(
        "Object o = \"z\"; String out = o instanceof String s ? s : \"n\"; System.out.println(out);",
    );
    assert_eq!(out, vec!["z"]);
}

#[test]
fn instanceof_pattern_in_ternary_else_branch() {
    let out = run_main(
        "Object o = 1; String out = o instanceof String s ? s : \"n\"; System.out.println(out);",
    );
    assert_eq!(out, vec!["n"]);
}

#[test]
fn instanceof_negated_in_if_condition() {
    let out = run_main(
        "Object o = 2; if (!(o instanceof String)) { System.out.println(\"not-string\"); }",
    );
    assert_eq!(out, vec!["not-string"]);
}

#[test]
fn instanceof_compound_with_starts_with_guard() {
    let out = run_main(
        "Object o = \"prefix-value\"; if (o instanceof String s && s.startsWith(\"prefix\")) { System.out.println(1); } else { System.out.println(0); }",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn instanceof_compound_with_equals_guard() {
    let out = run_main(
        "Object o = \"same\"; if (o instanceof String s && s.equals(\"same\")) { System.out.println(\"eq\"); } else { System.out.println(\"ne\"); }",
    );
    assert_eq!(out, vec!["eq"]);
}

#[test]
fn instanceof_double_pattern_binding() {
    let out = run_main(
        "Object o = 1.5; if (o instanceof Double d) { System.out.println(d.intValue()); }",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn instanceof_long_pattern_binding() {
    let out = run_main("Object o = 20L; if (o instanceof Long l) { System.out.println(l - 5L); }");
    assert_eq!(out, vec!["15"]);
}

#[test]
fn instanceof_boolean_pattern_binding() {
    let out = run_main("Object o = false; if (o instanceof Boolean b) { System.out.println(b); }");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn instanceof_character_pattern_binding() {
    let out = run_main("Object o = 'Q'; if (o instanceof Character c) { System.out.println(c); }");
    assert_eq!(out, vec!["Q"]);
}

#[test]
fn instanceof_nested_if_pattern_chain() {
    let out = run_main(
        "Object o = \"ab\"; if (o instanceof String s) { if (s.length() == 2) { System.out.println(\"pair\"); } }",
    );
    assert_eq!(out, vec!["pair"]);
}

#[test]
fn instanceof_pattern_else_if_chain() {
    let out = run_main(
        "Object o = 3; if (o instanceof String s) { System.out.println(\"s\"); } else if (o instanceof Integer n) { System.out.println(n); } else { System.out.println(\"x\"); }",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn instanceof_while_condition_with_string_pattern() {
    let out = run_main(
        "Object o = \"aa\"; int count = 0; while (o instanceof String s && count < s.length()) { count++; } System.out.println(count);",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn instanceof_exact_runtime_class_match() {
    let types = r#"
        static class Base {}
        static class Mid extends Base {}
        static class Leaf extends Mid {}
    "#;
    let out = run_in_main(
        "Base b = new Leaf(); System.out.println(b instanceof Leaf); System.out.println(b instanceof Mid);",
        types,
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn instanceof_pattern_with_logical_not_on_bound_check() {
    let out = run_main(
        "Object o = \"\"; if (o instanceof String s && !s.isEmpty()) { System.out.println(1); } else { System.out.println(0); }",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn instanceof_pattern_record_runtime_type() {
    let types = r#"
        record Pair(int a, int b) {}
    "#;
    let out = run_in_main(
        "Object o = new Pair(2, 3); if (o instanceof Pair p) { System.out.println(p.a() + p.b()); }",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn instanceof_compound_or_short_circuits_on_first_match() {
    let out = run_main(
        "Object o = \"z\"; boolean hit = o instanceof Integer || o instanceof String; System.out.println(hit);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn instanceof_compound_and_short_circuits_on_failed_type() {
    let out = run_main(
        "Object o = 1; boolean hit = o instanceof String s && s.length() > 0; System.out.println(hit);",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn instanceof_pattern_on_wrapper_short_type() {
    let out =
        run_main("Object o = (short) 6; if (o instanceof Short s) { System.out.println(s); }");
    assert_eq!(out, vec!["6"]);
}

#[test]
fn instanceof_pattern_on_wrapper_byte_type() {
    let out = run_main("Object o = (byte) 11; if (o instanceof Byte b) { System.out.println(b); }");
    assert_eq!(out, vec!["11"]);
}

#[test]
fn instanceof_negation_with_null_is_true() {
    let out = run_main("Object o = null; System.out.println(!(o instanceof String));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn instanceof_pattern_in_return_expression_via_helper() {
    let types = r#"
        static String label(Object o) {
            if (o instanceof String s) { return s; }
            return "other";
        }
    "#;
    let out = run_in_main(
        "System.out.println(label(\"tag\")); System.out.println(label(1));",
        types,
    );
    assert_eq!(out, vec!["tag", "other"]);
}

#[test]
fn instanceof_compound_three_part_condition() {
    let out = run_main(
        "Object o = \"abc\"; if (o instanceof String s && s.length() > 1 && s.charAt(0) == 'a') { System.out.println(\"match\"); } else { System.out.println(\"no\"); }",
    );
    assert_eq!(out, vec!["match"]);
}

#[test]
fn instanceof_pattern_before_explicit_cast_alternative() {
    let types = r#"
        static class Node { int id = 5; }
    "#;
    let out = run_in_main(
        "Object o = new Node(); if (o instanceof Node n) { System.out.println(n.id); } else { Node n = (Node) o; System.out.println(n.id); }",
        types,
    );
    assert_eq!(out, vec!["5"]);
}
