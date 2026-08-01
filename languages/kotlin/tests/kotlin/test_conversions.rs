use crate::helpers::run_prints;

#[test]
fn test_numeric_to_string_roundtrip() {
    let out = run_prints(
        r#"
        fun main() {
            val n = 15
            val fromNumber = n.toString()
            val parsed = fromNumber.toInt()
            println(fromNumber)
            println(parsed)
            println(parsed == n)
        }
    "#,
    );
    assert_eq!(out, &["15", "15", "true"]);
}

#[test]
fn test_double_to_int_truncation() {
    let out = run_prints(
        r#"
        fun main() {
            println(9.9.toInt())
            println((-9.9).toInt())
        }
    "#,
    );
    assert_eq!(out, &["9", "-9"]);
}

#[test]
fn test_int_to_double_and_long_arithmetic() {
    let out = run_prints(
        r#"
        fun main() {
            val value: Int = 7
            val asDouble = value.toDouble()
            val doubled = asDouble * 2.0
            println(doubled)
            println(doubled.toInt())
        }
    "#,
    );
    assert_eq!(out, &["14", "14"]);
}

#[test]
fn test_boolean_to_string_and_parse() {
    let out = run_prints(
        r#"
        fun main() {
            val flag = true
            val text = flag.toString()
            val truthy = "true".toBoolean()
            val falsy = "false".toBoolean()
            println(text)
            println(truthy)
            println(falsy)
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "false"]);
}

#[test]
fn test_character_codepoint_roundtrip() {
    let out = run_prints(
        r#"
        fun main() {
            val ch = 'A'
            val text = ch.toString()
            println(text)
            println(ch)
        }
    "#,
    );
    assert_eq!(out, &["A", "A"]);
}

#[test]
fn test_string_to_double_roundtrip() {
    let out = run_prints(
        r#"
        fun main() {
            val source = "3.75"
            val value = source.toDouble()
            println(value)
            println(value.toString())
        }
    "#,
    );
    assert_eq!(out, &["3.75", "3.75"]);
}

#[test]
fn test_string_to_int_failure_path_candidate() {
    let out = run_prints(
        r#"
        fun tryParseOrDefault(value: String): Int {
            return try {
                value.toInt()
            } catch (e: Exception) {
                -1
            }
        }

        fun main() {
            println(tryParseOrDefault("7"))
            println(tryParseOrDefault("bad"))
        }
    "#,
    );
    assert_eq!(out, &["7", "-1"]);
}

#[test]
fn test_nullable_to_string_behavior() {
    let out = run_prints(
        r#"
        fun main() {
            val value: String? = null
            println(value == null)
            val fallback = value?.toString() ?: "none"
            println(fallback)
        }
    "#,
    );
    assert_eq!(out, &["true", "none"]);
}

#[test]
fn test_string_to_int_handles_sign_prefixes() {
    let out = run_prints(
        r#"
        fun main() {
            println("+12".toInt())
            println("-12".toInt())
            println("0".toInt())
        }
    "#,
    );
    assert_eq!(out, &["12", "-12", "0"]);
}

#[test]
fn test_string_to_int_rejects_whitespace() {
    let out = run_prints(
        r#"
        fun safe(value: String): Int {
            return try {
                value.toInt()
            } catch (e: Exception) {
                -1
            }
        }

        fun main() {
            println(safe(" 12"))
            println(safe("12 "))
            println(safe("\t12"))
        }
    "#,
    );
    assert_eq!(out, &["-1", "-1", "-1"]);
}

#[test]
fn test_string_to_long_roundtrip() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 1234567890123L
            val text = value.toString()
            val parsed = text.toLong()
            println(text)
            println(parsed == value)
        }
    "#,
    );
    assert_eq!(out, &["1234567890123", "true"]);
}

#[test]
fn test_string_to_long_and_negative_roundtrip() {
    let out = run_prints(
        r#"
        fun main() {
            val value = -987654321L
            val parsed = value.toString().toLong()
            println(parsed)
            println(parsed == value)
        }
    "#,
    );
    assert_eq!(out, &["-987654321", "true"]);
}

#[test]
fn test_int_to_long_roundtrip() {
    let out = run_prints(
        r#"
        fun main() {
            val source: Int = 42
            val widened = source.toLong()
            val narrowed = widened.toInt()
            println(widened)
            println(narrowed)
            println(narrowed == source)
        }
    "#,
    );
    assert_eq!(out, &["42", "42", "true"]);
}

#[test]
fn test_long_to_int_overflow_wraps() {
    let out = run_prints(
        r#"
        fun main() {
            val large = 3_000_000_000L
            println(large.toInt())
        }
    "#,
    );
    assert_eq!(out, &["-1294967296"]);
}

#[test]
fn test_double_to_int_and_long_behavior() {
    let out = run_prints(
        r#"
        fun main() {
            val source = 12.9
            println(source.toInt())
            println(source.toLong())
        }
    "#,
    );
    assert_eq!(out, &["12", "12"]);
}

#[test]
fn test_double_negative_truncation_edges() {
    let out = run_prints(
        r#"
        fun main() {
            println((-9.0).toInt())
            println((-9.1).toInt())
            println((-9.9).toInt())
            println((-0.9).toInt())
        }
    "#,
    );
    assert_eq!(out, &["-9", "-9", "-9", "0"]);
}

#[test]
fn test_float_conversion_roundtrip_candidate() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 3.75f
            val asString = value.toString()
            val asDouble = asString.toDouble()
            println(asString)
            println(asDouble)
            println(asDouble == 3.75)
        }
    "#,
    );
    assert_eq!(out, &["3.75", "3.75", "true"]);
}

#[test]
fn test_string_to_double_scientific_notation() {
    let out = run_prints(
        r#"
        fun main() {
            println("1.5e2".toDouble())
            println("2.5E-1".toDouble())
        }
    "#,
    );
    assert_eq!(out, &["150", "0.25"]);
}

#[test]
fn test_string_to_double_rejects_invalid_and_whitespace() {
    let out = run_prints(
        r#"
        fun safe(value: String): Double {
            return try {
                value.toDouble()
            } catch (e: Exception) {
                Double.NaN
            }
        }

        fun main() {
            println(safe("nan"))
            println(safe(" 2.0"))
            println(safe("bad"))
        }
    "#,
    );
    assert_eq!(out, &["NaN", "NaN", "NaN"]);
}

#[test]
fn test_boolean_case_insensitive_parse() {
    let out = run_prints(
        r#"
        fun main() {
            println("TRUE".toBoolean())
            println("False".toBoolean())
            println("fAlSe".toBoolean())
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "false"]);
}

#[test]
fn test_boolean_to_string_and_nullable_to_boolean() {
    let out = run_prints(
        r#"
        fun main() {
            val a = false.toString()
            val b = true.toString()
            val maybe: Boolean? = null
            println(a)
            println(b)
            println(maybe?.toString() ?: "null")
        }
    "#,
    );
    assert_eq!(out, &["false", "true", "null"]);
}

#[test]
fn test_to_byte_wraps_lower_bit_slice() {
    let out = run_prints(
        r#"
        fun main() {
            println(127.toByte().toInt())
            println(128.toByte().toInt())
            println(255.toByte().toInt())
            println((-129).toByte().toInt())
        }
    "#,
    );
    assert_eq!(out, &["127", "-128", "-1", "127"]);
}

#[test]
fn test_to_short_wraps_lower_bits() {
    let out = run_prints(
        r#"
        fun main() {
            println(32767.toShort().toInt())
            println(32768.toShort().toInt())
            println((-32768).toShort().toInt())
            println((-32769).toShort().toInt())
        }
    "#,
    );
    assert_eq!(out, &["32767", "-32768", "-32768", "32767"]);
}

#[test]
fn test_float_to_int_rounding_floor_like_behavior() {
    let out = run_prints(
        r#"
        fun main() {
            println(9.999999.toInt())
            println((-9.999999).toInt())
        }
    "#,
    );
    assert_eq!(out, &["9", "-9"]);
}

#[test]
fn test_zero_and_negative_zero_to_string() {
    let out = run_prints(
        r#"
        fun main() {
            println(0.toString())
            println((-0).toDouble().toInt())
            println((-0.0).toString())
        }
    "#,
    );
    assert_eq!(out, &["0", "0", "0"]);
}

#[test]
fn test_character_code_conversions() {
    let out = run_prints(
        r#"
        fun main() {
            val text = "AZ"
            val first = text[0]
            val second = text[1]
            val asString = first.toString() + second.toString()
            println(asString)
            println(first.code)
            println(second.code)
        }
    "#,
    );
    assert_eq!(out, &["AZ", "65", "90"]);
}

#[test]
fn test_string_to_int_default_candidate() {
    let out = run_prints(
        r#"
        fun parseOrDefault(source: String, fallback: Int): Int {
            return try {
                source.toInt()
            } catch (e: Exception) {
                fallback
            }
        }

        fun main() {
            println(parseOrDefault("18", 7))
            println(parseOrDefault("bad", 7))
        }
    "#,
    );
    assert_eq!(out, &["18", "7"]);
}

#[test]
fn test_string_to_double_default_candidate() {
    let out = run_prints(
        r#"
        fun parseOrDefaultDouble(source: String, fallback: Double): Double {
            return try {
                source.toDouble()
            } catch (e: Exception) {
                fallback
            }
        }

        fun main() {
            println(parseOrDefaultDouble("5.5", 9.5))
            println(parseOrDefaultDouble("bad", 9.5))
        }
    "#,
    );
    assert_eq!(out, &["5.5", "9.5"]);
}

#[test]
fn test_plus_prefixed_numeric_strings_convert() {
    let out = run_prints(
        r#"
        fun main() {
            println("+42".toInt())
            println("+3.25".toDouble())
            println("0007".toInt())
        }
    "#,
    );
    assert_eq!(out, &["42", "3.25", "7"]);
}

#[test]
fn test_byte_to_int_extension_roundtrip() {
    let out = run_prints(
        r#"
        fun main() {
            val source: Byte = 13
            val text = source.toString()
            val round = text.toInt()
            println(source)
            println(round)
            println(round.toByte())
        }
    "#,
    );
    assert_eq!(out, &["13", "13", "13"]);
}

#[test]
fn test_string_to_int_radix_variants() {
    let out = run_prints(
        r#"
        fun main() {
            println("1010".toInt(2))
            println("ff".toInt(16))
            println("-11".toInt(3))
            println("77".toInt(8))
        }
    "#,
    );
    assert_eq!(out, &["10", "255", "-4", "63"]);
}

#[test]
fn test_string_to_int_radix_invalid_char_throws() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                println("2".toInt(2))
            } catch (e: NumberFormatException) {
                println("bad")
            }
        }
    "#,
    );
    assert_eq!(out, &["bad"]);
}

#[test]
fn test_string_to_long_radix_negative_and_large() {
    let out = run_prints(
        r#"
        fun main() {
            println("7fffffff".toLong(16))
            println("-100000000".toLong(2))
            println("1fffffffffffff".toLong(16))
        }
    "#,
    );
    assert_eq!(out, &["2147483647", "-256", "9007199254740991"]);
}

#[test]
fn test_character_codepoint_and_reverse_conversion() {
    let out = run_prints(
        r#"
        fun main() {
            val ch = 'Ω'
            val code = ch.code
            println(code)
            println(code.toChar())
            println(ch.toString())
        }
    "#,
    );
    assert_eq!(out, &["937", "Ω", "Ω"]);
}

#[test]
fn test_double_parse_infinite_and_nan_keywords() {
    let out = run_prints(
        r#"
        fun main() {
            println("Infinity".toDouble())
            println("-Infinity".toDouble())
            println("NaN".toDouble().isNaN())
        }
    "#,
    );
    assert_eq!(out, &["Infinity", "-Infinity", "true"]);
}

#[test]
fn test_boolean_parse_with_numeric_aliases_is_false() {
    let out = run_prints(
        r#"
        fun main() {
            println("1".toBoolean())
            println("0".toBoolean())
            println("TRUE ".toBoolean())
            println(" false ".toBoolean())
        }
    "#,
    );
    assert_eq!(out, &["false", "false", "false", "false"]);
}

#[test]
fn test_number_to_string_preserves_roundtrip_for_float_and_double() {
    let out = run_prints(
        r#"
        fun main() {
            val a = 1.0f
            val b = 2.5f
            val d = 3.5
            val fromStringToFloat = a.toString().toFloat()
            val fromStringToDouble = d.toString().toDouble()
            println(a.toString())
            println(fromStringToFloat)
            println(fromStringToDouble)
            println(b.toString())
            println(d.toString())
        }
    "#,
    );
    assert_eq!(out, &["1.0", "1.0", "3.5", "2.5", "3.5"]);
}

#[test]
fn test_short_to_byte_roundtrip_and_overflow_boundaries() {
    let out = run_prints(
        r#"
        fun main() {
            val a: Short = 32000
            val asByte = a.toByte()
            val restored = asByte.toShort()
            println(a.toByte().toInt())
            println(restored)
            println((-129).toByte().toInt())
        }
    "#,
    );
    assert_eq!(out, &["-96", "-96", "127"]);
}

#[test]
fn test_string_to_int_or_null_and_numeric_nullability() {
    let out = run_prints(
        r#"
        fun main() {
            println("42".toIntOrNull() ?: -1)
            println("nope".toIntOrNull() ?: -1)
            println("-9".toIntOrNull() ?: 0)
        }
    "#,
    );
    assert_eq!(out, &["42", "-1", "-9"]);
}

#[test]
fn test_string_to_double_or_null_rejects_bad_and_nan_variants() {
    let out = run_prints(
        r#"
        fun main() {
            println("3.5".toDoubleOrNull() ?: -1.0)
            println("bad".toDoubleOrNull() ?: -1.0)
            println("NaN".toDoubleOrNull()?.isNaN() ?: false)
        }
    "#,
    );
    assert_eq!(out, &["3.5", "-1", "true"]);
}

#[test]
fn test_string_to_long_or_null_radix_boundary() {
    let out = run_prints(
        r#"
        fun main() {
            println("7fffffff".toLongOrNull(16) ?: 0)
            println("xyz".toLongOrNull(16) ?: -1)
        }
    "#,
    );
    assert_eq!(out, &["2147483647", "-1"]);
}

#[test]
fn test_character_to_int_and_to_string_is_roundtrip() {
    let out = run_prints(
        r#"
        fun main() {
            val ch = 'z'
            val code = ch.code
            val decoded = code.toChar()
            println(code)
            println(decoded)
        }
    "#,
    );
    assert_eq!(out, &["122", "z"]);
}

#[test]
fn test_to_boolean_or_null_distinguishes_truthy_and_garbage() {
    let out = run_prints(
        r#"
        fun main() {
            println("true".toBooleanOrNull() ?: "null")
            println("FALSE".toBooleanOrNull() ?: "null")
            println("yes".toBooleanOrNull() ?: "null")
            println("0".toBooleanOrNull() ?: "null")
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "null", "null"]);
}
