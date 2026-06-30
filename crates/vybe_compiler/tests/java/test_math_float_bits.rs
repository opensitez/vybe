use crate::helpers::run_main;

#[test]
fn math_scalb_doubles_exponent() {
    let out = run_main(
        "System.out.println((int) Math.scalb(2.0, 2));",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn math_scalb_halves_with_negative_exp() {
    let out = run_main(
        "System.out.println((int) Math.scalb(8.0, -1));",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn math_ulp_of_one() {
    let out = run_main(
        "System.out.println(Math.ulp(1.0) > 0.0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn math_ulp_of_zero() {
    let out = run_main(
        "System.out.println(Math.ulp(0.0) > 0.0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn math_get_exponent_of_eight() {
    let out = run_main(
        "System.out.println(Math.getExponent(8.0));",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn math_get_exponent_of_one() {
    let out = run_main(
        "System.out.println(Math.getExponent(1.0));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn math_copy_sign_positive() {
    let out = run_main(
        "System.out.println((int) Math.copySign(5.0, -1.0));",
    );
    assert_eq!(out, vec!["-5"]);
}

#[test]
fn math_copy_sign_negative() {
    let out = run_main(
        "System.out.println((int) Math.copySign(-5.0, 1.0));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn math_next_after_toward_smaller() {
    let out = run_main(
        "System.out.println(Math.nextAfter(1.0, 0.0) < 1.0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn math_next_after_toward_larger() {
    let out = run_main(
        "System.out.println(Math.nextAfter(1.0, 2.0) > 1.0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn math_next_up_from_one() {
    let out = run_main(
        "System.out.println(Math.nextUp(1.0) > 1.0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn math_next_down_from_one() {
    let out = run_main(
        "System.out.println(Math.nextDown(1.0) < 1.0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn math_fma_multiply_add() {
    let out = run_main(
        "System.out.println((int) Math.fma(2.0, 3.0, 4.0));",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn math_fma_with_negative() {
    let out = run_main(
        "System.out.println((int) Math.fma(2.0, 3.0, -1.0));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn math_signum_positive() {
    let out = run_main(
        "System.out.println((int) Math.signum(5.0));",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn math_signum_negative() {
    let out = run_main(
        "System.out.println((int) Math.signum(-5.0));",
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn math_signum_zero() {
    let out = run_main(
        "System.out.println((int) Math.signum(0.0));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn math_rint_rounds_half() {
    let out = run_main(
        "System.out.println((int) Math.rint(2.5));",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn math_cbrt_of_eight() {
    let out = run_main(
        "System.out.println((int) Math.cbrt(8.0));",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn math_hypot_three_four() {
    let out = run_main(
        "System.out.println((int) Math.hypot(3.0, 4.0));",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn math_expm1_small() {
    let out = run_main(
        "System.out.println((int) Math.expm1(0.0));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn math_log1p_zero() {
    let out = run_main(
        "System.out.println((int) Math.log1p(0.0));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn math_to_degrees_half_pi() {
    let out = run_main(
        "System.out.println((int) Math.toDegrees(Math.PI / 2));",
    );
    assert_eq!(out, vec!["90"]);
}

#[test]
fn math_to_radians_180() {
    let out = run_main(
        "System.out.println((int) Math.toRadians(180.0));",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn math_ieee_remainder() {
    let out = run_main(
        "System.out.println((int) Math.IEEEremainder(7.0, 4.0));",
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn math_increment_exact_int() {
    let out = run_main(
        "System.out.println(Math.incrementExact(5));",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn math_decrement_exact_int() {
    let out = run_main(
        "System.out.println(Math.decrementExact(5));",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn math_negate_exact_int() {
    let out = run_main(
        "System.out.println(Math.negateExact(7));",
    );
    assert_eq!(out, vec!["-7"]);
}

#[test]
fn math_floor_div_int() {
    let out = run_main(
        "System.out.println(Math.floorDiv(7, 2));",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn math_floor_mod_int() {
    let out = run_main(
        "System.out.println(Math.floorMod(7, 2));",
    );
    assert_eq!(out, vec!["1"]);
}

