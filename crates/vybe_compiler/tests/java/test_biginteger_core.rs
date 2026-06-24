use crate::helpers::run_main;

#[test]
fn biginteger_add_two_positive_values() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("123"); java.math.BigInteger b = new java.math.BigInteger("456"); System.out.println(a.add(b).toString());"#,
    );
    assert_eq!(out, vec!["579"]);
}

#[test]
fn biginteger_add_negative_and_positive() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("-50"); java.math.BigInteger b = new java.math.BigInteger("80"); System.out.println(a.add(b).toString());"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn biginteger_subtract_yields_difference() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("100"); java.math.BigInteger b = new java.math.BigInteger("37"); System.out.println(a.subtract(b).toString());"#,
    );
    assert_eq!(out, vec!["63"]);
}

#[test]
fn biginteger_subtract_negative_result() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("5"); java.math.BigInteger b = new java.math.BigInteger("12"); System.out.println(a.subtract(b).toString());"#,
    );
    assert_eq!(out, vec!["-7"]);
}

#[test]
fn biginteger_multiply_small_factors() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("6"); java.math.BigInteger b = new java.math.BigInteger("7"); System.out.println(a.multiply(b).toString());"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn biginteger_multiply_by_zero_yields_zero() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("999"); System.out.println(a.multiply(java.math.BigInteger.ZERO).toString());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn biginteger_mod_returns_remainder() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("29"); java.math.BigInteger b = new java.math.BigInteger("5"); System.out.println(a.mod(b).toString());"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn biginteger_mod_with_larger_divisor() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("7"); java.math.BigInteger b = new java.math.BigInteger("11"); System.out.println(a.mod(b).toString());"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn biginteger_gcd_of_coprime_numbers_is_one() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("8"); java.math.BigInteger b = new java.math.BigInteger("15"); System.out.println(a.gcd(b).toString());"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn biginteger_gcd_of_multiples_finds_common_factor() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("54"); java.math.BigInteger b = new java.math.BigInteger("24"); System.out.println(a.gcd(b).toString());"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn biginteger_pow_small_exponent_squares() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("3"); System.out.println(a.pow(2).toString());"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn biginteger_pow_cube() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("5"); System.out.println(a.pow(3).toString());"#,
    );
    assert_eq!(out, vec!["125"]);
}

#[test]
fn biginteger_pow_zero_exponent_yields_one() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("99"); System.out.println(a.pow(0).toString());"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn biginteger_pow_one_exponent_is_identity() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("42"); System.out.println(a.pow(1).toString());"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn biginteger_value_of_long_literal() {
    let out = run_main(
        r#"java.math.BigInteger a = java.math.BigInteger.valueOf(12345L); System.out.println(a.toString());"#,
    );
    assert_eq!(out, vec!["12345"]);
}

#[test]
fn biginteger_value_of_negative_long() {
    let out = run_main(
        r#"java.math.BigInteger a = java.math.BigInteger.valueOf(-99L); System.out.println(a.toString());"#,
    );
    assert_eq!(out, vec!["-99"]);
}

#[test]
fn biginteger_compare_to_equal_values_is_zero() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("100"); java.math.BigInteger b = new java.math.BigInteger("100"); System.out.println(a.compareTo(b));"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn biginteger_compare_to_smaller_is_negative() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("3"); java.math.BigInteger b = new java.math.BigInteger("10"); System.out.println(a.compareTo(b));"#,
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn biginteger_compare_to_larger_is_positive() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("50"); java.math.BigInteger b = new java.math.BigInteger("7"); System.out.println(a.compareTo(b));"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn biginteger_zero_constant_value() {
    let out = run_main(
        r#"java.math.BigInteger a = java.math.BigInteger.ZERO; System.out.println(a.toString());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn biginteger_one_constant_value() {
    let out = run_main(
        r#"java.math.BigInteger a = java.math.BigInteger.ONE; System.out.println(a.toString());"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn biginteger_ten_constant_value() {
    let out = run_main(
        r#"java.math.BigInteger a = java.math.BigInteger.TEN; System.out.println(a.multiply(a).toString());"#,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn biginteger_negate_flips_sign() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("17"); System.out.println(a.negate().toString());"#,
    );
    assert_eq!(out, vec!["-17"]);
}

#[test]
fn biginteger_abs_of_negative_value() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("-42"); System.out.println(a.abs().toString());"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn biginteger_signum_positive_value() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("1"); System.out.println(a.signum());"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn biginteger_signum_negative_value() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("-1"); System.out.println(a.signum());"#,
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn biginteger_signum_zero_value() {
    let out = run_main(
        r#"java.math.BigInteger a = java.math.BigInteger.ZERO; System.out.println(a.signum());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn biginteger_max_picks_larger_operand() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("4"); java.math.BigInteger b = new java.math.BigInteger("9"); System.out.println(a.max(b).toString());"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn biginteger_min_picks_smaller_operand() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("4"); java.math.BigInteger b = new java.math.BigInteger("9"); System.out.println(a.min(b).toString());"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn biginteger_bit_length_of_small_value() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("8"); System.out.println(a.bitLength());"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn biginteger_test_bit_checks_set_bit() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("5"); System.out.println(a.testBit(0)); System.out.println(a.testBit(2));"#,
    );
    assert_eq!(out, vec!["true", "true"]);
}

#[test]
fn biginteger_shift_left_multiplies_by_power_of_two() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("3"); System.out.println(a.shiftLeft(2).toString());"#,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn biginteger_shift_right_divides_by_power_of_two() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("20"); System.out.println(a.shiftRight(2).toString());"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn biginteger_and_bitwise_masks_bits() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("15"); java.math.BigInteger b = new java.math.BigInteger("10"); System.out.println(a.and(b).toString());"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn biginteger_or_bitwise_combines_bits() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("12"); java.math.BigInteger b = new java.math.BigInteger("3"); System.out.println(a.or(b).toString());"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn biginteger_xor_bitwise_exclusive_or() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("15"); java.math.BigInteger b = new java.math.BigInteger("10"); System.out.println(a.xor(b).toString());"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn biginteger_not_bitwise_complement() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("0"); System.out.println(a.not().toString());"#,
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn biginteger_is_probable_prime_on_small_prime() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("17"); System.out.println(a.isProbablePrime(10));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn biginteger_next_probable_prime_advances_from_even() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("14"); System.out.println(a.nextProbablePrime().toString());"#,
    );
    assert_eq!(out, vec!["17"]);
}

#[test]
fn biginteger_add_large_decimal_string_values() {
    let out = run_main(
        r#"java.math.BigInteger a = new java.math.BigInteger("999999999999999999"); java.math.BigInteger b = new java.math.BigInteger("1"); System.out.println(a.add(b).toString());"#,
    );
    assert_eq!(out, vec!["1000000000000000000"]);
}
