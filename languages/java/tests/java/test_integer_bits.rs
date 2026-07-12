use crate::helpers::run_main;

#[test]
fn integer_bit_count_zero_has_no_one_bits() {
    let out = run_main("System.out.println(Integer.bitCount(0));");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn integer_bit_count_one_has_single_one_bit() {
    let out = run_main("System.out.println(Integer.bitCount(1));");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn integer_bit_count_three_ones_in_value_seven() {
    let out = run_main("System.out.println(Integer.bitCount(7));");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn integer_bit_count_counts_bits_in_ten() {
    let out = run_main("System.out.println(Integer.bitCount(10));");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn integer_bit_count_all_bits_set_in_negative_one() {
    let out = run_main("System.out.println(Integer.bitCount(-1));");
    assert_eq!(out, vec!["32"]);
}

#[test]
fn integer_bit_count_power_of_two_has_one_bit() {
    let out = run_main("System.out.println(Integer.bitCount(256));");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn integer_highest_one_bit_zero_returns_zero() {
    let out = run_main("System.out.println(Integer.highestOneBit(0));");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn integer_highest_one_bit_one_returns_one() {
    let out = run_main("System.out.println(Integer.highestOneBit(1));");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn integer_highest_one_bit_five_returns_four() {
    let out = run_main("System.out.println(Integer.highestOneBit(5));");
    assert_eq!(out, vec!["4"]);
}

#[test]
fn integer_highest_one_bit_ten_returns_eight() {
    let out = run_main("System.out.println(Integer.highestOneBit(10));");
    assert_eq!(out, vec!["8"]);
}

#[test]
fn integer_highest_one_bit_negative_value_uses_sign_bit() {
    let out = run_main("System.out.println(Integer.highestOneBit(-5));");
    assert_eq!(out, vec!["-2147483648"]);
}

#[test]
fn integer_highest_one_bit_max_value_returns_itself() {
    let out = run_main("System.out.println(Integer.highestOneBit(Integer.MAX_VALUE));");
    assert_eq!(out, vec!["1073741824"]);
}

#[test]
fn integer_lowest_one_bit_zero_returns_zero() {
    let out = run_main("System.out.println(Integer.lowestOneBit(0));");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn integer_lowest_one_bit_six_returns_two() {
    let out = run_main("System.out.println(Integer.lowestOneBit(6));");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn integer_lowest_one_bit_eight_returns_eight() {
    let out = run_main("System.out.println(Integer.lowestOneBit(8));");
    assert_eq!(out, vec!["8"]);
}

#[test]
fn integer_lowest_one_bit_twelve_returns_four() {
    let out = run_main("System.out.println(Integer.lowestOneBit(12));");
    assert_eq!(out, vec!["4"]);
}

#[test]
fn integer_lowest_one_bit_negative_six_returns_two() {
    let out = run_main("System.out.println(Integer.lowestOneBit(-6));");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn integer_lowest_one_bit_one_returns_one() {
    let out = run_main("System.out.println(Integer.lowestOneBit(1));");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn integer_rotate_left_one_by_one_yields_two() {
    let out = run_main("System.out.println(Integer.rotateLeft(1, 1));");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn integer_rotate_left_five_by_two_yields_twenty() {
    let out = run_main("System.out.println(Integer.rotateLeft(5, 2));");
    assert_eq!(out, vec!["20"]);
}

#[test]
fn integer_rotate_left_high_bit_wraps_to_one() {
    let out = run_main("System.out.println(Integer.rotateLeft(0x80000000, 1));");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn integer_rotate_left_zero_distance_is_identity() {
    let out = run_main("System.out.println(Integer.rotateLeft(42, 0));");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn integer_rotate_left_full_word_is_identity() {
    let out = run_main("System.out.println(Integer.rotateLeft(12345, 32));");
    assert_eq!(out, vec!["12345"]);
}

#[test]
fn integer_rotate_left_negative_distance_rotates_right() {
    let out = run_main("System.out.println(Integer.rotateLeft(8, -1));");
    assert_eq!(out, vec!["4"]);
}

#[test]
fn integer_number_of_leading_zeros_zero_is_thirty_two() {
    let out = run_main("System.out.println(Integer.numberOfLeadingZeros(0));");
    assert_eq!(out, vec!["32"]);
}

#[test]
fn integer_number_of_leading_zeros_one_is_thirty_one() {
    let out = run_main("System.out.println(Integer.numberOfLeadingZeros(1));");
    assert_eq!(out, vec!["31"]);
}

#[test]
fn integer_number_of_leading_zeros_high_bit_set_is_zero() {
    let out = run_main("System.out.println(Integer.numberOfLeadingZeros(0x80000000));");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn integer_number_of_leading_zeros_eight_is_twenty_eight() {
    let out = run_main("System.out.println(Integer.numberOfLeadingZeros(8));");
    assert_eq!(out, vec!["28"]);
}

#[test]
fn integer_number_of_leading_zeros_max_value_is_zero() {
    let out = run_main("System.out.println(Integer.numberOfLeadingZeros(Integer.MAX_VALUE));");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn integer_number_of_leading_zeros_negative_one_is_zero() {
    let out = run_main("System.out.println(Integer.numberOfLeadingZeros(-1));");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn integer_to_binary_string_ten_is_one_zero_one_zero() {
    let out = run_main(r#"System.out.println(Integer.toBinaryString(10));"#);
    assert_eq!(out, vec!["1010"]);
}

#[test]
fn integer_to_binary_string_zero_is_single_zero() {
    let out = run_main(r#"System.out.println(Integer.toBinaryString(0));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn integer_to_binary_string_one_is_single_one() {
    let out = run_main(r#"System.out.println(Integer.toBinaryString(1));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn integer_to_binary_string_negative_one_is_all_ones() {
    let out = run_main(r#"System.out.println(Integer.toBinaryString(-1));"#);
    assert_eq!(out, vec!["11111111111111111111111111111111"]);
}

#[test]
fn integer_to_binary_string_eight_is_one_followed_by_zeros() {
    let out = run_main(r#"System.out.println(Integer.toBinaryString(8));"#);
    assert_eq!(out, vec!["1000"]);
}

#[test]
fn integer_to_hex_string_two_fifty_five_is_ff() {
    let out = run_main(r#"System.out.println(Integer.toHexString(255));"#);
    assert_eq!(out, vec!["ff"]);
}

#[test]
fn integer_to_hex_string_sixteen_is_ten() {
    let out = run_main(r#"System.out.println(Integer.toHexString(16));"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn integer_to_hex_string_zero_is_zero() {
    let out = run_main(r#"System.out.println(Integer.toHexString(0));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn integer_to_hex_string_ten_is_a() {
    let out = run_main(r#"System.out.println(Integer.toHexString(10));"#);
    assert_eq!(out, vec!["a"]);
}

#[test]
fn integer_to_hex_string_negative_one_is_all_f_digits() {
    let out = run_main(r#"System.out.println(Integer.toHexString(-1));"#);
    assert_eq!(out, vec!["ffffffff"]);
}
