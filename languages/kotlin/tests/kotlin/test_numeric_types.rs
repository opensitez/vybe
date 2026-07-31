use crate::helpers::run_prints;

#[test]
fn test_integer_addition_and_subtraction() {
    let out = run_prints(r#"
        fun main() {
            println(10 + 5)
            println(10 - 5)
            println(-3 + 7)
            println(3 - 10)
        }
    "#);
    assert_eq!(out, &["15", "5", "4", "-7"]);
}

#[test]
fn test_integer_multiplication_and_zero_identity() {
    let out = run_prints(r#"
        fun main() {
            println(7 * 6)
            println(7 * 0)
            println(0 * 9)
            println(-3 * 4)
        }
    "#);
    assert_eq!(out, &["42", "0", "0", "-12"]);
}

#[test]
fn test_integer_division_truncates_toward_zero() {
    let out = run_prints(r#"
        fun main() {
            println(7 / 3)
            println(8 / 4)
            println(-7 / 3)
            println(7 / -3)
        }
    "#);
    assert_eq!(out, &["2", "2", "-2", "-2"]);
}

#[test]
fn test_integer_remainder_follows_dividend_sign() {
    let out = run_prints(r#"
        fun main() {
            println(7 % 3)
            println(7 % 4)
            println(-7 % 3)
            println(7 % -4)
        }
    "#);
    assert_eq!(out, &["1", "3", "-1", "3"]);
}

#[test]
fn test_integer_division_by_zero_is_runtime_error() {
    let out = run_prints(r#"
        fun main() {
            try {
                println(7 / 0)
            } catch (e: Exception) {
                println("division-error")
            }
        }
    "#);
    assert_eq!(out, &["division-error"]);
}

#[test]
fn test_float_division_keeps_fractional_part() {
    let out = run_prints(r#"
        fun main() {
            println(7.0 / 2.0)
            println(1.0 / 2.0)
            println(-7.0 / 2.0)
        }
    "#);
    assert_eq!(out, &["3.5", "0.5", "-3.5"]);
}

#[test]
fn test_remainder_with_floating_values() {
    let out = run_prints(r#"
        fun main() {
            println(7.5 % 2.0)
            println(8.2 % 2.0)
            println(-7.5 % 2.0)
        }
    "#);
    assert_eq!(out, &["1.5", "0.2", "-1.5"]);
}

#[test]
fn test_long_basic_arithmetic() {
    let out = run_prints(r#"
        fun main() {
            val a: Long = 1_000_000_000_000
            val b: Long = 250
            println(a + b)
            println(a - b)
            println(a * 2)
            println(a / b)
        }
    "#);
    assert_eq!(out, &["1000000000250", "999999999750", "2000000000000", "4000000"]);
}

#[test]
fn test_mixed_int_and_long_math_prefers_wide_type() {
    let out = run_prints(r#"
        fun main() {
            val base = 3
            val wide = 10L
            println(base + wide)
            println(base * wide)
            println(wide - base)
            println(wide / base)
        }
    "#);
    assert_eq!(out, &["13", "30", "7", "3"]);
}

#[test]
fn test_int_and_double_mix_rounds_through_double() {
    let out = run_prints(r#"
        fun main() {
            val value = 5
            println(value + 2.5)
            println(value * 1.5)
            println(value / 2.0)
            println(10 / 4 + 0.5)
        }
    "#);
    assert_eq!(out, &["7.5", "7.5", "2.5", "3.5"]);
}

#[test]
fn test_arithmetic_precedence_is_standard() {
    let out = run_prints(r#"
        fun main() {
            println(2 + 3 * 4)
            println((2 + 3) * 4)
            println(10 - 6 / 2 + 3)
            println(10 - (6 / (2 + 1)))
        }
    "#);
    assert_eq!(out, &["14", "20", "10", "8"]);
}

#[test]
fn test_unary_plus_and_unary_minus() {
    let out = run_prints(r#"
        fun main() {
            val a = 7
            val b = -7
            println(+a)
            println(-a)
            println(+b)
            println(-b)
        }
    "#);
    assert_eq!(out, &["7", "-7", "-7", "7"]);
}

#[test]
fn test_increment_prefix_and_postfix() {
    let out = run_prints(r#"
        fun main() {
            var count = 5
            println(++count)
            println(count)
            println(count++)
            println(count)
        }
    "#);
    assert_eq!(out, &["6", "6", "6", "7"]);
}

#[test]
fn test_decrement_prefix_and_postfix() {
    let out = run_prints(r#"
        fun main() {
            var value = 10
            println(--value)
            println(value)
            println(value--)
            println(value)
        }
    "#);
    assert_eq!(out, &["9", "9", "9", "8"]);
}

#[test]
fn test_compound_plus_assign() {
    let out = run_prints(r#"
        fun main() {
            var total = 1
            total += 4
            total += 5
            println(total)
            total += -2
            println(total)
        }
    "#);
    assert_eq!(out, &["10", "8"]);
}

#[test]
fn test_compound_minus_assign() {
    let out = run_prints(r#"
        fun main() {
            var total = 20
            total -= 3
            total -= 5
            println(total)
            total -= -2
            println(total)
        }
    "#);
    assert_eq!(out, &["12", "14"]);
}

#[test]
fn test_compound_times_assign() {
    let out = run_prints(r#"
        fun main() {
            var total = 2
            total *= 3
            total *= 4
            println(total)
            total *= -1
            println(total)
        }
    "#);
    assert_eq!(out, &["24", "-24"]);
}

#[test]
fn test_compound_divide_assign_integer() {
    let out = run_prints(r#"
        fun main() {
            var total = 32
            total /= 4
            total /= 2
            println(total)
        }
    "#);
    assert_eq!(out, &["4"]);
}

#[test]
fn test_compound_mod_assign() {
    let out = run_prints(r#"
        fun main() {
            var value = 23
            value %= 5
            println(value)
            value %= 2
            println(value)
        }
    "#);
    assert_eq!(out, &["3", "1"]);
}

#[test]
fn test_bitwise_and_masking() {
    let out = run_prints(r#"
        fun main() {
            val value = 0b1111 and 0b1010
            val mask = 0b0101 and value
            println(value)
            println(mask)
        }
    "#);
    assert_eq!(out, &["10", "0"]);
}

#[test]
fn test_bitwise_or_and_xor() {
    let out = run_prints(r#"
        fun main() {
            val a = 0b1110
            val b = 0b0110
            println(a or b)
            println(a xor b)
        }
    "#);
    assert_eq!(out, &["14", "8"]);
}

#[test]
fn test_bitwise_not_inverts_bits() {
    let out = run_prints(r#"
        fun main() {
            val value = 0
            println(value.inv())
            println((-1).inv())
        }
    "#);
    assert_eq!(out, &["-1", "0"]);
}

#[test]
fn test_bit_shift_left_and_right() {
    let out = run_prints(r#"
        fun main() {
            val value = 1
            println(value shl 4)
            println(16 shr 2)
            println(-16 shr 2)
        }
    "#);
    assert_eq!(out, &["16", "4", "-4"]);
}

#[test]
fn test_unsigned_shift_right() {
    let out = run_prints(r#"
        fun main() {
            println(-1 ushr 1)
            println(-16 ushr 2)
            println(16 ushr 1)
        }
    "#);
    assert_eq!(out, &["2147483647", "1073741820", "8"]);
}

#[test]
fn test_int_min_and_max_boundaries_are_stable() {
    let out = run_prints(r#"
        fun main() {
            println(Int.MAX_VALUE)
            println(Int.MIN_VALUE)
            println(Long.MAX_VALUE)
            println(Long.MIN_VALUE)
        }
    "#);
    assert_eq!(out, &["2147483647", "-2147483648", "9223372036854775807", "-9223372036854775808"]);
}

#[test]
fn test_int_boundary_arithmetic_wraps_with_two_complement() {
    let out = run_prints(r#"
        fun main() {
            println(Int.MAX_VALUE + 1)
            println(Int.MIN_VALUE - 1)
            println(Long.MAX_VALUE + 1)
            println(Long.MIN_VALUE - 1)
        }
    "#);
    assert_eq!(out, &["-2147483648", "2147483647", "-9223372036854775808", "9223372036854775807"]);
}

#[test]
fn test_numeric_increment_and_branching() {
    let out = run_prints(r#"
        fun main() {
            var step = 0
            while (step < 4) {
                step++
            }
            println(step)
            var done = step > 3
            println(done)
        }
    "#);
    assert_eq!(out, &["4", "true"]);
}

#[test]
fn test_long_and_int_comparison() {
    let out = run_prints(r#"
        fun main() {
            val intValue = 100
            val longValue = 100L
            println(intValue == longValue)
            println(intValue < longValue + 1)
            println((intValue + 1) >= longValue)
        }
    "#);
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_float_to_int_and_to_long_truncate_toward_zero() {
    let out = run_prints(r#"
        fun main() {
            println(2.9.toInt())
            println(-2.9.toInt())
            println(2.9.toLong())
            println(-2.9.toLong())
        }
    "#);
    assert_eq!(out, &["2", "-2", "2", "-2"]);
}

#[test]
fn test_byte_and_short_roundtrip_via_int() {
    let out = run_prints(r#"
        fun main() {
            val b: Byte = 127
            val s: Short = 32767
            println(b.toInt() + 1)
            println(s.toInt() + 1)
            println(b.toLong() - 7)
            println(s.toLong() - 7)
        }
    "#);
    assert_eq!(out, &["128", "32768", "120", "32760"]);
}

#[test]
fn test_numeric_relational_operator_network() {
    let out = run_prints(r#"
        fun main() {
            println(3 > 2)
            println(3 >= 3)
            println(3 < 4.0)
            println(3L != 4L)
            println(3.0 == 3)
        }
    "#);
    assert_eq!(out, &["true", "true", "true", "true", "true"]);
}

#[test]
fn test_numeric_zero_rules() {
    let out = run_prints(r#"
        fun main() {
            println(0)
            println(-0)
            println(0 + 0)
            println(0L)
            println(0.0)
            println(-0.0)
            println(0.0 == -0.0)
        }
    "#);
    assert_eq!(out, &["0", "0", "0", "0", "0", "0", "true"]);
}
