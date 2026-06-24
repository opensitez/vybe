use crate::helpers::run_main;

#[test]
fn integer_modulo_remainder() {
    let out = run_main("System.out.println(17 % 5);");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn unary_negation_on_variable() {
    let out = run_main("int x = 8; System.out.println(-x);");
    assert_eq!(out, vec!["-8"]);
}

#[test]
fn bitwise_and_masks_bits() {
    let out = run_main("System.out.println(0b1100 & 0b1010);");
    assert_eq!(out, vec!["8"]);
}

#[test]
fn logical_and_requires_both_true() {
    let out = run_main(
        "boolean a = true; boolean b = false; System.out.println(a && b); System.out.println(a && true);",
    );
    assert_eq!(out, vec!["false", "true"]);
}

#[test]
fn shift_left_doubles_value() {
    let out = run_main("System.out.println(3 << 2);");
    assert_eq!(out, vec!["12"]);
}

#[test]
fn logical_or_accepts_either_true_operand() {
    let out = run_main("System.out.println(false || true);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn equality_compares_primitive_values() {
    let out = run_main("System.out.println(5 == 5); System.out.println(5 == 6);");
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn relational_operators_order_integers() {
    let out = run_main(
        "System.out.println(2 < 3); System.out.println(3 <= 3); System.out.println(4 > 5);",
    );
    assert_eq!(out, vec!["true", "true", "false"]);
}

#[test]
fn bitwise_xor_sets_bits_different_in_operands() {
    let out = run_main("System.out.println(0b1100 ^ 0b1010);");
    assert_eq!(out, vec!["6"]);
}
