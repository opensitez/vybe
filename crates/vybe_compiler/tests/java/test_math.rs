use crate::helpers::run_main;

#[test]
fn math_abs_on_negative_int() {
    let out = run_main("System.out.println(Math.abs(-15));");
    assert_eq!(out, vec!["15"]);
}

#[test]
fn math_max_picks_larger_argument() {
    let out = run_main("System.out.println(Math.max(3, 9));");
    assert_eq!(out, vec!["9"]);
}

#[test]
fn math_sqrt_of_perfect_square() {
    let out = run_main("System.out.println((int) Math.sqrt(81));");
    assert_eq!(out, vec!["9"]);
}

#[test]
fn math_round_half_up_for_positive() {
    let out = run_main("System.out.println(Math.round(2.6));");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn integer_compare_to_orders_values() {
    let out = run_main(
        "System.out.println(Integer.compare(5, 8)); System.out.println(Integer.compare(8, 5));",
    );
    assert_eq!(out, vec!["-1", "1"]);
}
