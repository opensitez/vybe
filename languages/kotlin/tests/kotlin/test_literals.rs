use crate::helpers::run_prints;

#[test]
fn test_integer_literals_decimal_and_negative() {
    let out = run_prints(
        r#"
        fun main() {
            println(0)
            println(123)
            println(-456)
        }
    "#,
    );
    assert_eq!(out, &["0", "123", "-456"]);
}

#[test]
fn test_integer_literals_hex_and_binary() {
    let out = run_prints(
        r#"
        fun main() {
            println(0x1A)
            println(0x10 + 1)
            println(0b1010)
            println(0b1111 + 1)
        }
    "#,
    );
    assert_eq!(out, &["26", "17", "10", "16"]);
}

#[test]
fn test_integer_literal_underscores() {
    let out = run_prints(
        r#"
        fun main() {
            println(1_000)
            println(10_000_000)
            println(1_2_3_4)
        }
    "#,
    );
    assert_eq!(out, &["1000", "10000000", "1234"]);
}

#[test]
fn test_long_integer_literals() {
    let out = run_prints(
        r#"
        fun main() {
            val short = 10L
            val big = 1_000_000_000L
            val neg = -12L
            println(short)
            println(big)
            println(neg)
            println(short + neg)
        }
    "#,
    );
    assert_eq!(out, &["10", "1000000000", "-12", "-2"]);
}

#[test]
fn test_boolean_literals() {
    let out = run_prints(
        r#"
        fun main() {
            val a = true
            val b = false
            println(a && !b)
            println(a || b)
            println(!a && b)
            println(true == true)
            println(false == false)
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "false", "true", "true"]);
}

#[test]
fn test_floating_literals_basic() {
    let out = run_prints(
        r#"
        fun main() {
            println(1.0)
            println(0.5)
            println(-2.25)
            println(3.14)
        }
    "#,
    );
    assert_eq!(out, &["1", "0.5", "-2.25", "3.14"]);
}

#[test]
fn test_floating_literals_scientific_notation() {
    let out = run_prints(
        r#"
        fun main() {
            println(1e3)
            println(2.5e-1)
            println(3.0E2)
            println(1e-3)
        }
    "#,
    );
    assert_eq!(out, &["1000", "0.25", "300", "0.001"]);
}

#[test]
fn test_float_suffix_literal() {
    let out = run_prints(
        r#"
        fun main() {
            val tiny: Float = 1.25f
            val rounded = tiny + 0.75f
            println(tiny)
            println(rounded)
            println(rounded.toString())
        }
    "#,
    );
    assert_eq!(out, &["1.25", "2.0", "2.0"]);
}

#[test]
fn test_double_suffix_literal() {
    let out = run_prints(
        r#"
        fun main() {
            val pi = 3.1415
            val alsoPi = 3.1415d
            println(pi)
            println(alsoPi)
            println(alsoPi == pi)
        }
    "#,
    );
    assert_eq!(out, &["3.1415", "3.1415", "true"]);
}

#[test]
fn test_character_literals_plain_and_escaped() {
    let out = run_prints(
        r#"
        fun main() {
            val letter = 'K'
            val quote = '\''
            val backslash = '\\'
            println(letter)
            println(quote)
            println(backslash)
        }
    "#,
    );
    assert_eq!(out, &["K", "'", "\\"]);
}

#[test]
fn test_character_unicode_literal() {
    let out = run_prints(
        r#"
        fun main() {
            val omega = '\u03A9'
            val heart = '\u2665'
            println(omega)
            println(heart)
        }
    "#,
    );
    assert_eq!(out, &["Ω", "♥"]);
}

#[test]
fn test_string_basic_literals() {
    let out = run_prints(
        r#"
        fun main() {
            println("plain")
            println("")
            println("x")
        }
    "#,
    );
    assert_eq!(out, &["plain", "", "x"]);
}

#[test]
fn test_string_escape_sequences() {
    let out = run_prints(
        r#"
        fun main() {
            println("a\tb")
            println("a\nb")
            println("\"quoted\"")
            println("c\\d")
        }
    "#,
    );
    assert_eq!(out, &["a\tb", "a", "b", "\"quoted\"", "c\\d"]);
}

#[test]
fn test_string_template_with_expression() {
    let out = run_prints(
        r#"
        fun main() {
            val a = 2
            val b = 4
            println("$a + $b = ${a + b}")
        }
    "#,
    );
    assert_eq!(out, &["2 + 4 = 6"]);
}

#[test]
fn test_string_template_escapes_dollar_sign() {
    let out = run_prints(
        r#"
        fun main() {
            println("${'$'}4.99")
            println("${'$'}{a + 2}")
            val prefix = "\${prefix}"
            println(prefix)
        }
    "#,
    );
    assert_eq!(out, &["$4.99", "${a + 2}", "${prefix}"]);
}

#[test]
fn test_multiline_raw_string_basic() {
    let out = run_prints(
        r#"
        fun main() {
            val block = """
line one
line two
"""
            println(block.trimIndent())
            println(block.lines().size)
        }
    "#,
    );
    assert_eq!(out, &["line one", "line two", "2"]);
}

#[test]
fn test_raw_string_with_margin() {
    let out = run_prints(
        r#"
        fun main() {
            val block = """
                >a
                >b
            """.trimMargin(">")
            println(block)
            println(block.lines().size)
        }
    "#,
    );
    assert_eq!(out, &["a", "b", "2"]);
}

#[test]
fn test_null_literal_and_type_inference() {
    let out = run_prints(
        r#"
        fun main() {
            val value: String? = null
            val safe = value ?: "none"
            println(value == null)
            println(safe)
            val explicit: String = "ok"
            println(explicit)
        }
    "#,
    );
    assert_eq!(out, &["true", "none", "ok"]);
}

#[test]
fn test_mixed_literal_types_in_tuple() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(1, 2.0, true, 'x')
            println(values.size)
            println(values[0])
            println(values[1])
            println(values[2])
            println(values[3])
        }
    "#,
    );
    assert_eq!(out, &["4", "1", "2.0", "true", "x"]);
}

#[test]
fn test_collection_from_literal_values() {
    let out = run_prints(
        r#"
        fun main() {
            println(listOf(1, 2, 3).joinToString(","))
            println(listOf("a", "b", "c").size)
            println(arrayOf(1.0, 2.5).joinToString("|"))
            println(intArrayOf(1, 2, 3).size)
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "3", "1.0|2.5", "3"]);
}

#[test]
fn test_boolean_logic_literal_interactions() {
    let out = run_prints(
        r#"
        fun main() {
            println(true && false)
            println(false || true)
            println(!false)
            println(true && (1 > 0))
        }
    "#,
    );
    assert_eq!(out, &["false", "true", "true", "true"]);
}

#[test]
fn test_parenthesized_and_unary_literals() {
    let out = run_prints(
        r#"
        fun main() {
            println(-(1 + 2))
            println(-(2 * 3))
            println(+5)
            println(+(-5))
            println((-1L) + 2L)
        }
    "#,
    );
    assert_eq!(out, &["-3", "-6", "5", "-5", "1"]);
}

#[test]
fn test_empty_and_null_coalescing_literal_source() {
    let out = run_prints(
        r#"
        fun main() {
            val nullable: String? = null
            println(nullable?.let { it.length } ?: -1)
            val present: String? = "k"
            println(present?.length ?: 0)
        }
    "#,
    );
    assert_eq!(out, &["-1", "1"]);
}

#[test]
fn test_multiline_comment_style_string_content() {
    let out = run_prints(
        r#"
        fun main() {
            val raw = """
                /*
                one
                */
                """.trimIndent()
            println(raw.contains("/*"))
            println(raw.lines().size)
        }
    "#,
    );
    assert_eq!(out, &["true", "3"]);
}

#[test]
fn test_short_and_byte_numeric_literals() {
    let out = run_prints(
        r#"
        fun main() {
            val small: Short = 12
            val tiny: Byte = 7
            val unsigned: Int = 1
            println(small)
            println(tiny)
            println(unsigned)
            println(small + tiny + unsigned)
        }
    "#,
    );
    assert_eq!(out, &["12", "7", "1", "20"]);
}

#[test]
fn test_boolean_literal_in_control_flow() {
    let out = run_prints(
        r#"
        fun main() {
            val yes = true
            val no = false
            val chosen = if (yes && !no) "go" else "stop"
            val result = if (yes && no) "bad" else if (!no) "ok" else "never"
            println(chosen)
            println(result)
        }
    "#,
    );
    assert_eq!(out, &["go", "ok"]);
}

#[test]
fn test_char_and_string_literal_mixing() {
    let out = run_prints(
        r#"
        fun main() {
            val c: Char = 'x'
            val text = "c=$c"
            val pair = listOf(c, 'y')
            println(text)
            println(pair.joinToString("-"))
        }
    "#,
    );
    assert_eq!(out, &["c=x", "x-y"]);
}

#[test]
fn test_floating_nan_and_infinity_literals() {
    let out = run_prints(
        r#"
        fun main() {
            val zero = 0.0
            val nan = 0.0 / zero
            val inf = 1.0 / zero
            println(nan.isNaN())
            println(inf > 0)
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_literal_infix_chaining_readability() {
    let out = run_prints(
        r#"
        fun main() {
            val value = 1 + 2 * 3 - 4 / 2
            val grouped = (1 + 2) * (3 - 4 / 2)
            val withFloats = 1 + 2.0 * 2.5 - 3.0
            println(value)
            println(grouped)
            println(withFloats)
        }
    "#,
    );
    assert_eq!(out, &["5", "3", "3.0"]);
}

#[test]
fn test_zero_length_and_empty_char_array_literals() {
    let out = run_prints(
        r#"
        fun main() {
            val empty = ""
            val letters = charArrayOf()
            println(empty.isEmpty())
            println(letters.isEmpty())
            println(empty == "")
            println(letters.size)
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true", "0"]);
}
