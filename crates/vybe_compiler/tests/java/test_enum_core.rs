use crate::helpers::{run_in_main, run_main};

#[test]
fn enum_constants_print_declared_names() {
    let types = r#"
        enum Season { SPRING, SUMMER, FALL, WINTER }
    "#;
    let out = run_in_main(
        "System.out.println(Season.SPRING); System.out.println(Season.WINTER);",
        types,
    );
    assert_eq!(out, vec!["SPRING", "WINTER"]);
}

#[test]
fn enum_first_constant_has_ordinal_zero() {
    let types = r#"
        enum Rank { BRONZE, SILVER, GOLD }
    "#;
    let out = run_in_main("System.out.println(Rank.BRONZE.ordinal());", types);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn enum_last_constant_ordinal_matches_index() {
    let types = r#"
        enum Level { LOW, MID, HIGH }
    "#;
    let out = run_in_main("System.out.println(Level.HIGH.ordinal());", types);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn enum_middle_constant_ordinal_is_one() {
    let types = r#"
        enum Axis { X, Y, Z }
    "#;
    let out = run_in_main("System.out.println(Axis.Y.ordinal());", types);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn enum_value_of_resolves_constant_by_name() {
    let types = r#"
        enum Mode { ON, OFF }
    "#;
    let out = run_in_main("System.out.println(Mode.valueOf(\"ON\"));", types);
    assert_eq!(out, vec!["ON"]);
}

#[test]
fn enum_value_of_returns_second_declared_constant() {
    let types = r#"
        enum Pair { LEFT, RIGHT }
    "#;
    let out = run_in_main("System.out.println(Pair.valueOf(\"RIGHT\"));", types);
    assert_eq!(out, vec!["RIGHT"]);
}

#[test]
fn enum_switch_matches_declared_constant() {
    let types = r#"
        enum Hue { RED, GREEN, BLUE }
    "#;
    let out = run_in_main(
        "Hue h = Hue.GREEN; switch (h) { case RED: System.out.println(\"r\"); break; case GREEN: System.out.println(\"g\"); break; default: System.out.println(\"b\"); }",
        types,
    );
    assert_eq!(out, vec!["g"]);
}

#[test]
fn enum_switch_hits_default_for_unlisted_constant() {
    let types = r#"
        enum Token { A, B, C }
    "#;
    let out = run_in_main(
        "Token t = Token.C; switch (t) { case A: System.out.println(\"a\"); break; case B: System.out.println(\"b\"); break; default: System.out.println(\"other\"); }",
        types,
    );
    assert_eq!(out, vec!["other"]);
}

#[test]
fn enum_switch_first_case_without_break_falls_through() {
    let types = r#"
        enum Step { ONE, TWO, THREE }
    "#;
    let out = run_in_main(
        "Step s = Step.ONE; switch (s) { case ONE: System.out.println(\"1\"); case TWO: System.out.println(\"2\"); break; default: System.out.println(\"x\"); }",
        types,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn enum_switch_break_stops_further_cases() {
    let types = r#"
        enum Flag { YES, NO }
    "#;
    let out = run_in_main(
        "Flag f = Flag.YES; switch (f) { case YES: System.out.println(\"y\"); break; case NO: System.out.println(\"n\"); }",
        types,
    );
    assert_eq!(out, vec!["y"]);
}

#[test]
fn enum_field_stores_constructor_argument() {
    let types = r#"
        enum Coin { PENNY(1), NICKEL(5);
            final int cents;
            Coin(int c) { cents = c; }
        }
    "#;
    let out = run_in_main(
        "System.out.println(Coin.PENNY.cents); System.out.println(Coin.NICKEL.cents);",
        types,
    );
    assert_eq!(out, vec!["1", "5"]);
}

#[test]
fn enum_field_assigns_distinct_constructor_values() {
    let types = r#"
        enum Code { A(10), B(20), C(30);
            final int code;
            Code(int c) { code = c; }
        }
    "#;
    let out = run_in_main(
        "System.out.println(Code.A.code); System.out.println(Code.C.code);",
        types,
    );
    assert_eq!(out, vec!["10", "30"]);
}

#[test]
fn enum_instance_method_returns_custom_label() {
    let types = r#"
        enum Shape { CIRCLE, SQUARE;
            String label() { return \"shape\"; }
        }
    "#;
    let out = run_in_main("System.out.println(Shape.CIRCLE.label());", types);
    assert_eq!(out, vec!["shape"]);
}

#[test]
fn enum_method_reads_own_field_value() {
    let types = r#"
        enum Unit { KB(1024), MB(1048576);
            final int bytes;
            Unit(int b) { bytes = b; }
            int kilobytes() { return bytes / 1024; }
        }
    "#;
    let out = run_in_main("System.out.println(Unit.MB.kilobytes());", types);
    assert_eq!(out, vec!["1024"]);
}

#[test]
fn enum_equality_same_constant_is_true() {
    let types = r#"
        enum State { OPEN, CLOSED }
    "#;
    let out = run_in_main(
        "State a = State.OPEN; State b = State.OPEN; System.out.println(a == b);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_inequality_different_constants_is_true() {
    let types = r#"
        enum Side { LEFT, RIGHT }
    "#;
    let out = run_in_main("System.out.println(Side.LEFT != Side.RIGHT);", types);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_reference_reassigned_to_another_constant() {
    let types = r#"
        enum Dir { N, S, E, W }
    "#;
    let out = run_in_main("Dir d = Dir.N; d = Dir.W; System.out.println(d);", types);
    assert_eq!(out, vec!["W"]);
}

#[test]
fn enum_switch_arrow_rule_prints_matched_label() {
    let types = r#"
        enum Grade { A, B, C }
    "#;
    let out = run_in_main(
        "Grade g = Grade.B; switch (g) { case A -> System.out.println(\"top\"); case B -> System.out.println(\"mid\"); default -> System.out.println(\"low\"); }",
        types,
    );
    assert_eq!(out, vec!["mid"]);
}

#[test]
fn enum_switch_expression_yields_string_label() {
    let types = r#"
        enum Size { S, M, L }
    "#;
    let out = run_in_main(
        "Size s = Size.L; String tag = switch (s) { case S -> \"small\"; case M -> \"medium\"; default -> \"large\"; }; System.out.println(tag);",
        types,
    );
    assert_eq!(out, vec!["large"]);
}

#[test]
fn enum_compare_ordinals_via_subtraction() {
    let types = r#"
        enum Phase { ALPHA, BETA, GA }
    "#;
    let out = run_in_main(
        "System.out.println(Phase.GA.ordinal() - Phase.ALPHA.ordinal());",
        types,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn enum_if_branch_selects_matching_constant() {
    let types = r#"
        enum Role { ADMIN, USER }
    "#;
    let out = run_in_main(
        "Role r = Role.ADMIN; if (r == Role.ADMIN) { System.out.println(\"admin\"); } else { System.out.println(\"user\"); }",
        types,
    );
    assert_eq!(out, vec!["admin"]);
}

#[test]
fn enum_two_member_type_supports_both_ordinals() {
    let types = r#"
        enum Bit { ZERO, ONE }
    "#;
    let out = run_in_main(
        "System.out.println(Bit.ZERO.ordinal()); System.out.println(Bit.ONE.ordinal());",
        types,
    );
    assert_eq!(out, vec!["0", "1"]);
}

#[test]
fn enum_string_field_exposes_payload() {
    let types = r#"
        enum Lang { JAVA(\"java\"), RUST(\"rust\");
            final String id;
            Lang(String id) { this.id = id; }
        }
    "#;
    let out = run_in_main(
        "System.out.println(Lang.JAVA.id); System.out.println(Lang.RUST.id);",
        types,
    );
    assert_eq!(out, vec!["java", "rust"]);
}

#[test]
fn enum_boolean_method_reports_active_state() {
    let types = r#"
        enum Power { ON, OFF;
            boolean isOn() { return this == ON; }
        }
    "#;
    let out = run_in_main(
        "System.out.println(Power.ON.isOn()); System.out.println(Power.OFF.isOn());",
        types,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn enum_static_style_switch_on_local_variable() {
    let types = r#"
        enum Op { ADD, SUB, MUL }
    "#;
    let out = run_in_main(
        "Op op = Op.MUL; int code = 0; switch (op) { case ADD: code = 1; break; case SUB: code = 2; break; case MUL: code = 3; break; } System.out.println(code);",
        types,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn enum_value_of_result_compares_equal_to_constant() {
    let types = r#"
        enum Key { ENTER, ESC }
    "#;
    let out = run_in_main(
        "System.out.println(Key.valueOf(\"ESC\") == Key.ESC);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn enum_switch_last_declared_constant() {
    let types = r#"
        enum Month { JAN, FEB, MAR }
    "#;
    let out = run_in_main(
        "Month m = Month.MAR; switch (m) { case JAN: System.out.println(1); break; case FEB: System.out.println(2); break; case MAR: System.out.println(3); break; }",
        types,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn enum_method_invoked_from_main_on_each_constant() {
    let types = r#"
        enum Parity { EVEN, ODD;
            int code() { return ordinal(); }
        }
    "#;
    let out = run_in_main(
        "System.out.println(Parity.EVEN.code()); System.out.println(Parity.ODD.code());",
        types,
    );
    assert_eq!(out, vec!["0", "1"]);
}

#[test]
fn enum_field_shared_across_constants_with_same_payload() {
    let types = r#"
        enum Dup { A(1), B(1);
            final int n;
            Dup(int n) { this.n = n; }
        }
    "#;
    let out = run_in_main(
        "System.out.println(Dup.A.n); System.out.println(Dup.B.n);",
        types,
    );
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn enum_switch_default_when_variable_holds_first_constant() {
    let types = r#"
        enum Choice { X, Y }
    "#;
    let out = run_in_main(
        "Choice c = Choice.X; switch (c) { default: System.out.println(\"d\"); case X: System.out.println(\"x\"); break; }",
        types,
    );
    assert_eq!(out, vec!["x"]);
}
