use crate::helpers::run_prints;

#[test]
fn test_and_with_positive_values() {
    let out = run_prints(
        r#"
        fun main() {
            println(0b1111 and 0b1010)
            println(0b0101 and 0b1010)
            println(12 and 10)
            println(4 and 3)
        }
    "#,
    );
    assert_eq!(out, &["10", "0", "8", "0"]);
}

#[test]
fn test_or_with_positive_values() {
    let out = run_prints(
        r#"
        fun main() {
            println(0b1111 or 0b1010)
            println(0b0001 or 0b0010)
            println(4 or 3)
            println(8 or 1)
        }
    "#,
    );
    assert_eq!(out, &["15", "3", "7", "9"]);
}

#[test]
fn test_xor_with_positive_values() {
    let out = run_prints(
        r#"
        fun main() {
            println(0b1111 xor 0b1010)
            println(0b1111 xor 0b1111)
            println(6 xor 3)
            println(0b1100 xor 0b1010)
        }
    "#,
    );
    assert_eq!(out, &["5", "0", "5", "6"]);
}

#[test]
fn test_inv_for_simple_values() {
    let out = run_prints(
        r#"
        fun main() {
            println(0.inv())
            println(1.inv())
            println(255.inv())
            println(1023.inv())
        }
    "#,
    );
    assert_eq!(out, &["-1", "0", "-256", "-1024"]);
}

#[test]
fn test_shift_left_basic() {
    let out = run_prints(
        r#"
        fun main() {
            println(1 shl 0)
            println(1 shl 1)
            println(1 shl 4)
            println(3 shl 2)
        }
    "#,
    );
    assert_eq!(out, &["1", "2", "16", "12"]);
}

#[test]
fn test_shift_right_arithmetic_positive() {
    let out = run_prints(
        r#"
        fun main() {
            println(16 shr 1)
            println(16 shr 4)
            println(16 shr 5)
            println(3 shr 1)
        }
    "#,
    );
    assert_eq!(out, &["8", "1", "0", "1"]);
}

#[test]
fn test_shift_right_arithmetic_negative() {
    let out = run_prints(
        r#"
        fun main() {
            println(-8 shr 1)
            println(-8 shr 2)
            println(-1 shr 3)
            println(-15 shr 1)
        }
    "#,
    );
    assert_eq!(out, &["-4", "-2", "-1", "-8"]);
}

#[test]
fn test_unsigned_shift_right_basic() {
    let out = run_prints(
        r#"
        fun main() {
            println(-1 ushr 1)
            println(-16 ushr 2)
            println(16 ushr 1)
            println(1 ushr 1)
        }
    "#,
    );
    assert_eq!(out, &["2147483647", "1073741820", "8", "0"]);
}

#[test]
fn test_bit_masking_even_and_odd() {
    let out = run_prints(
        r#"
        fun main() {
            val mask = 1
            val values = listOf(1, 2, 3, 4, 5, 6, 7, 8)
            val onlyEven = values.filter { it and 1 == 0 }
            val onlyOdd = values.filter { it and 1 == 1 }
            println(onlyEven.joinToString(","))
            println(onlyOdd.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,4,6,8", "1,3,5,7"]);
}

#[test]
fn test_nybble_high_and_low_flags() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 0b11010101
            println(value and 0b1111)
            println(value shr 4)
            println(value and 0b11110000)
            println((value and 0b11110000) shr 4)
        }
    "#,
    );
    assert_eq!(out, &["5", "13", "208", "13"]);
}

#[test]
fn test_flags_with_union_and_intersection() {
    let out = run_prints(
        r#"
        fun main() {
            val canRead = 0b001
            val canWrite = 0b010
            val canExecute = 0b100
            val perms = canRead or canWrite
            println(perms and canRead)
            println(perms and canExecute)
            val withExec = perms or canExecute
            val withoutRead = withExec and canRead.inv()
            println(withExec)
            println(withoutRead)
            println(withoutRead and canWrite)
            println(withoutRead and canExecute)
        }
    "#,
    );
    assert_eq!(out, &["1", "0", "7", "-8", "2", "4"]);
}

#[test]
fn test_long_bitwise_and() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Long = 0xFFL
            val mask: Long = 0x0F
            println(value and mask)
            println((value xor mask))
            println(value or mask)
        }
    "#,
    );
    assert_eq!(out, &["15", "240", "255"]);
}

#[test]
fn test_long_bitwise_or_and_xor() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Long = 0b1010
            val next: Long = 0b1100
            println(value or next)
            println(value and next)
            println(value xor next)
        }
    "#,
    );
    assert_eq!(out, &["14", "8", "6"]);
}

#[test]
fn test_long_shift_left() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Long = 1L
            println(value shl 8)
            println((1L shl 32).toString())
            println((3L shl 5))
        }
    "#,
    );
    assert_eq!(out, &["256", "4294967296", "96"]);
}

#[test]
fn test_long_shift_right_arithmetic() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Long = 64L
            val negative: Long = -64L
            println(value shr 2)
            println(negative shr 3)
            println(negative shr 2)
            println(negative shr 1)
        }
    "#,
    );
    assert_eq!(out, &["16", "-8", "-16", "-32"]);
}

#[test]
fn test_long_unsigned_shift_right() {
    let out = run_prints(
        r#"
        fun main() {
            val negative: Long = -1L
            val signed: Long = -16L
            println(negative ushr 1)
            println(signed ushr 2)
            println(15L ushr 1)
            println(1L ushr 1)
        }
    "#,
    );
    assert_eq!(
        out,
        &["9223372036854775807", "2305843009213693950", "7", "0"]
    );
}

#[test]
fn test_bitwise_precedence_with_arithmetic() {
    let out = run_prints(
        r#"
        fun main() {
            println(1 or 2 + 4 and 8)
            println((1 or 2) + (4 and 8))
            println(2 shl 3 + 1)
            println(2 shl (3 + 1))
        }
    "#,
    );
    assert_eq!(out, &["9", "3", "16", "32"]);
}

#[test]
fn test_bitwise_precedes_comparison_logic() {
    let out = run_prints(
        r#"
        fun main() {
            val raw = 0b1010
            println((raw and 2) == 2)
            println((raw and 1) == 1)
            println((raw or 1) > raw)
            println((raw xor 0) == raw)
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "true", "true"]);
}

#[test]
fn test_bitwise_filters_using_shifted_masks() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(0, 1, 2, 3, 4, 5, 6, 7, 8, 15, 16, 31)
            val maskedTwoBits = values.map { it and 0b11 }
            val flags = values.filter { (it and 0b1000) == 0b1000 }
            println(maskedTwoBits.joinToString(","))
            println(flags.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["0,1,2,3,0,1,2,3,0,3,0,3", "8,15"]);
}

#[test]
fn test_bitwise_toggle_odd_bits() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(1, 2, 3, 4, 5, 6, 7)
            val toggled = values.map { it xor 1 }
            println(toggled.joinToString(","))
            val restored = toggled.map { it xor 1 }
            println(restored.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["0,3,2,5,4,7,6", "1,2,3,4,5,6,7"]);
}

#[test]
fn test_bitwise_setting_and_clearing_bits() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 0
            val withBit2 = value or (1 shl 2)
            val withBit3 = withBit2 or (1 shl 3)
            val cleared = withBit3 and (1 shl 2).inv()
            println(withBit2)
            println(withBit3)
            println(cleared)
        }
    "#,
    );
    assert_eq!(out, &["4", "12", "8"]);
}

#[test]
fn test_bitwise_counting_subset_flags() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(0b1010, 0b1111, 0b1000, 0b0011)
            val countAnyHigh = values.count { it and 0b1000 != 0 }
            val countZeroLow = values.count { it and 1 == 0 }
            val countPairs = values.filter { (it and 0b0110) == 0b0010 }
            println(countAnyHigh)
            println(countZeroLow)
            println(countPairs.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["3", "3", "10"]);
}

#[test]
fn test_bitwise_clear_all_low_bits() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 0xFFFF
            val cleared = value and 0xFFF0
            val low = value and 0x000F
            println(cleared)
            println(low)
            println((cleared and low))
        }
    "#,
    );
    assert_eq!(out, &["65520", "15", "0"]);
}

#[test]
fn test_bitwise_power_of_two_checks() {
    let out = run_prints(
        r#"
        fun main() {
            val one = 1
            val two = 1 shl 1
            val three = 1 shl 2
            val eight = 1 shl 3
            println(two and two)
            println(three and one)
            println(eight and 4)
            println(eight and eight)
        }
    "#,
    );
    assert_eq!(out, &["2", "0", "0", "8"]);
}

#[test]
fn test_bitwise_roundtrip_with_mask() {
    let out = run_prints(
        r#"
        fun main() {
            val original = 0b10101010
            val mask = 0b11110000
            val hidden = original and mask
            val shown = original and mask.inv()
            val visible = (original and mask.inv())
            println(hidden)
            println(shown)
            println(visible)
            println(hidden + visible)
        }
    "#,
    );
    assert_eq!(out, &["160", "10", "10", "170"]);
}

#[test]
fn test_bitwise_identity_with_self_xor() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(0, 1, 2, 3, 255)
            val unchanged = values.map { it xor it }
            val back = values.map { (it xor 0) xor it }
            println(unchanged.joinToString(","))
            println(back.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["0,0,0,0,0", "0,1,2,3,255"]);
}

#[test]
fn test_bitwise_identity_with_self_and() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(0, 1, 2, 3, 255)
            val kept = values.map { it and it }
            val zeroed = values.map { it and 0 }
            println(kept.joinToString(","))
            println(zeroed.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["0,1,2,3,255", "0,0,0,0,0"]);
}

#[test]
fn test_bitwise_identity_with_self_or() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(0, 1, 2, 3, 255)
            val same = values.map { it or it }
            val plus = values.map { it or 0 }
            println(same.joinToString(","))
            println(plus.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["0,1,2,3,255", "0,1,2,3,255"]);
}

#[test]
fn test_bitwise_roundtrip_with_shift_and_or() {
    let out = run_prints(
        r#"
        fun main() {
            val original = 0b101010
            val shifted = original shl 2
            val restored = (shifted shr 2) or (original and 0)
            println(shifted)
            println(restored)
        }
    "#,
    );
    assert_eq!(out, &["168", "42"]);
}

#[test]
fn test_bitwise_parity_test_with_masking() {
    let out = run_prints(
        r#"
        fun main() {
            val numbers = listOf(10, 11, 12, 13, 14, 15)
            val parity = numbers.map { it and 1 }
            val even = numbers.filter { it and 1 == 0 }
            println(parity.joinToString(","))
            println(even.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["0,1,0,1,0,1", "10,12,14"]);
}

#[test]
fn test_shift_count_wraps_for_int() {
    let out = run_prints(
        r#"
        fun main() {
            println(1 shl 31)
            println(1 shl 32)
            println(1 shl 40)
            println(1 shl -1)
        }
    "#,
    );
    assert_eq!(out, &["-2147483648", "1", "256", "-2147483648"]);
}

#[test]
fn test_shift_count_wraps_for_long() {
    let out = run_prints(
        r#"
        fun main() {
            println(1L shl 63)
            println(1L shl 64)
            println(1L shl 65)
            println(1L shl -1)
        }
    "#,
    );
    assert_eq!(
        out,
        &["-9223372036854775808", "1", "2", "-9223372036854775808"]
    );
}

#[test]
fn test_unsigned_right_shift_of_negative_masks_with_and() {
    let out = run_prints(
        r#"
        fun main() {
            val signed = -8
            val unsigned = signed ushr 2
            println(unsigned)
            println(unsigned and 0x3FFFFFFF)
        }
    "#,
    );
    assert_eq!(out, &["1073741822", "2"]);
}

#[test]
fn test_masking_chain_preserves_expected_bits() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 0b10101111
            val lowNibble = value and 0x0F
            val upperNibble = (value and 0xF0) ushr 4
            println(lowNibble)
            println(upperNibble)
            println(((upperNibble shl 4) or lowNibble))
        }
    "#,
    );
    assert_eq!(out, &["15", "10", "175"]);
}

#[test]
fn test_set_clear_and_toggle_idempotent() {
    let out = run_prints(
        r#"
        fun main() {
            val base = 0b1001
            val set2 = base or (1 shl 1)
            val clear2 = set2 and (1 shl 1).inv()
            val toggle = base xor (1 shl 2)
            val toggledBack = toggle xor (1 shl 2)
            println(set2)
            println(clear2)
            println(toggle)
            println(toggledBack)
        }
    "#,
    );
    assert_eq!(out, &["11", "9", "13", "9"]);
}

#[test]
fn test_isolate_least_significant_set_bit() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 0b1011000
            val lsb = value and (-value)
            println(lsb)
            println((value and (value - 1)))
        }
    "#,
    );
    assert_eq!(out, &["8", "88"]);
}

#[test]
fn test_short_and_byte_are_extended_before_bitwise() {
    let out = run_prints(
        r#"
        fun main() {
            val signedByte: Byte = -1
            val signedShort: Short = -2
            val byteUnsigned = signedByte.toInt() and 0xFF
            val shortUnsigned = signedShort.toInt() and 0xFFFF
            val combined = (byteUnsigned and shortUnsigned)
            println(byteUnsigned)
            println(shortUnsigned)
            println(combined)
        }
    "#,
    );
    assert_eq!(out, &["255", "65534", "254"]);
}

#[test]
fn test_bitwise_with_java_long_bitcount_and_number_of_leading_zeros() {
    let out = run_prints(
        r#"
        fun main() {
            val sample = 0b0001_0010
            println(java.lang.Integer.bitCount(sample))
            println(java.lang.Integer.numberOfLeadingZeros(sample))
            println(java.lang.Integer.numberOfTrailingZeros(sample))
            println(java.lang.Integer.numberOfTrailingZeros(0))
        }
    "#,
    );
    assert_eq!(out, &["2", "28", "1", "32"]);
}

#[test]
fn test_long_to_int_truncation_after_masking() {
    let out = run_prints(
        r#"
        fun main() {
            val wide: Long = 0x1_0000_0000L
            val narrowed = (wide and 0xFFFF_FFFF).toInt()
            println(wide.toString())
            println(narrowed)
            println(wide.toInt())
        }
    "#,
    );
    assert_eq!(out, &["4294967296", "0", "0"]);
}

#[test]
fn test_bitwise_is_equivalent_between_inline_and_functional_calls() {
    let out = run_prints(
        r#"
        fun main() {
            val base = 0b11011001
            val andResult = base and 0x0F
            val andAlt = kotlin.math.floor(base.toDouble()).toInt() and 0x0F
            println(andResult)
            println(andAlt)
            val invAnd = base and (1 shl 4).inv()
            println(invAnd)
        }
    "#,
    );
    assert_eq!(out, &["25", "25", "201"]);
}
