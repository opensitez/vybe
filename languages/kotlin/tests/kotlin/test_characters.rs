use crate::helpers::run_prints;

#[test]
fn test_single_character_output() {
    let out = run_prints(
        r#"
        fun main() {
            println('a')
            println('Z')
        }
    "#,
    );
    assert_eq!(out, &["a", "Z"]);
}

#[test]
fn test_ascii_codepoint_properties() {
    let out = run_prints(
        r#"
        fun main() {
            println('A'.code)
            println('z'.code)
            println('0'.code)
        }
    "#,
    );
    assert_eq!(out, &["65", "122", "48"]);
}

#[test]
fn test_upper_and_lower_case_checks() {
    let out = run_prints(
        r#"
        fun main() {
            println('a'.isLowerCase())
            println('B'.isUpperCase())
            println('9'.isLowerCase())
            println('9'.isUpperCase())
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "false", "false"]);
}

#[test]
fn test_character_letter_or_digit_predicates() {
    let out = run_prints(
        r#"
        fun main() {
            println('5'.isDigit())
            println('X'.isLetter())
            println('X'.isLetterOrDigit())
            println('#'.isLetterOrDigit())
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true", "false"]);
}

#[test]
fn test_character_digit_validation() {
    let out = run_prints(
        r#"
        fun main() {
            val chars = listOf('1', '9', 'a', ' ')
            println(chars.all { it.isDigit() })
            println(chars.count { it.isDigit() })
            println(chars[0].digitToInt())
            println(chars[1].digitToInt())
            println(chars[2].digitToIntOrNull() == null)
        }
    "#,
    );
    assert_eq!(out, &["false", "2", "1", "9", "true"]);
}

#[test]
fn test_character_whitespace_checks() {
    let out = run_prints(
        r#"
        fun main() {
            println(' '.isWhitespace())
            println('\t'.isWhitespace())
            println('\n'.isWhitespace())
            println('a'.isWhitespace())
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true", "false"]);
}

#[test]
fn test_character_case_conversion() {
    let out = run_prints(
        r#"
        fun main() {
            println('k'.uppercaseChar())
            println('M'.lowercaseChar())
            println('ß'.uppercaseChar())
            println('ß'.lowercaseChar())
        }
    "#,
    );
    assert_eq!(out, &["K", "m", "SS", "ß"]);
}

#[test]
fn test_title_case_character_conversion() {
    let out = run_prints(
        r#"
        fun main() {
            println('a'.titlecaseChar())
            println('b'.titlecaseChar())
            println('1'.titlecaseChar())
        }
    "#,
    );
    assert_eq!(out, &["A", "B", "1"]);
}

#[test]
fn test_character_equality_and_compare_to() {
    let out = run_prints(
        r#"
        fun main() {
            println('k' == 'k')
            println('k' != 'K')
            println('a' < 'c')
            println('z'.compareTo('a'))
            println('a'.compareTo('z'))
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true", "25", "-25"]);
}

#[test]
fn test_character_range_membership() {
    let out = run_prints(
        r#"
        fun main() {
            println('b' in 'a'..'f')
            println('z' in 'a'..'f')
            println('5' in '0'..'9')
            println('g' in 'a'..'f')
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "true", "false"]);
}

#[test]
fn test_character_to_string_and_joining() {
    let out = run_prints(
        r#"
        fun main() {
            val chars = charArrayOf('a', 'b', 'c')
            println(chars.joinToString(","))
            println(chars[0].toString())
        }
    "#,
    );
    assert_eq!(out, &["a,b,c", "a"]);
}

#[test]
fn test_character_mutable_string_append() {
    let out = run_prints(
        r#"
        fun main() {
            var value = ""
            val a = 'A'
            value += a
            value += ':'
            value += 'B'
            println(value)
            println(value.length)
        }
    "#,
    );
    assert_eq!(out, &["A:B", "3"]);
}

#[test]
fn test_character_unicode_literal() {
    let out = run_prints(
        r#"
        fun main() {
            val copyright = '\u00A9'
            val omega = '\u03A9'
            println(copyright)
            println(omega)
        }
    "#,
    );
    assert_eq!(out, &["©", "Ω"]);
}

#[test]
fn test_character_iteration_from_string() {
    let out = run_prints(
        r#"
        fun main() {
            var out = ""
            for (c in "kotlin") {
                out += c
            }
            val first = "kotlin"[0]
            val last = "kotlin"["kotlin".lastIndex]
            println(out)
            println(first)
            println(last)
        }
    "#,
    );
    assert_eq!(out, &["kotlin", "k", "n"]);
}

#[test]
fn test_character_index_of_in_string() {
    let out = run_prints(
        r#"
        fun main() {
            val word = "banana"
            println(word.indexOf('a'))
            println(word.indexOf('a', 2))
            println(word.lastIndexOf('a'))
            println(word.count { it == 'a' })
        }
    "#,
    );
    assert_eq!(out, &["1", "3", "5", "3"]);
}

#[test]
fn test_character_filter_with_unicode_category() {
    let out = run_prints(
        r#"
        fun main() {
            val input = "a1b2c3"
            val letters = input.filter { it.isLetter() }
            val digits = input.filter { it.isDigit() }
            val mapped = input.map { if (it.isDigit()) '*' else it }
            println(letters)
            println(digits)
            println(mapped.joinToString(""))
        }
    "#,
    );
    assert_eq!(out, &["abc", "123", "a*b*c*"]);
}

#[test]
fn test_character_to_upper_lower_roundtrip() {
    let out = run_prints(
        r#"
        fun main() {
            println('x'.uppercaseChar().lowercaseChar())
            println('X'.lowercaseChar().uppercaseChar())
            println('ß'.lowercaseChar())
            println('ß'.uppercaseChar())
        }
    "#,
    );
    assert_eq!(out, &["x", "X", "ß", "SS"]);
}

#[test]
fn test_character_is_control() {
    let out = run_prints(
        r#"
        fun main() {
            println('\u0000'.isISOControl())
            println('\n'.isISOControl())
            println('a'.isISOControl())
            println('\u0009'.isISOControl())
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "false", "true"]);
}

#[test]
fn test_character_array_contains_and_any_all() {
    let out = run_prints(
        r#"
        fun main() {
            val chars = charArrayOf('k', 'o', 't', 'l', 'i', 'n')
            println(chars.contains('t'))
            println(chars.contains('a'))
            println(chars.any { it.isVowel() })
            println(chars.all { it.isLetter() })
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "true", "true"]);
}

#[test]
fn test_character_combine_case_predicate() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "aB3!"
            val lowered = value.map { it.lowercaseChar() }.joinToString("")
            val uppered = value.map { it.uppercaseChar() }.joinToString("")
            println(lowered)
            println(uppered)
            println(lowered[1].isUpperCase())
            println(uppered[2].isDigit())
        }
    "#,
    );
    assert_eq!(out, &["ab3!", "AB3!", "false", "true"]);
}

#[test]
fn test_character_count_in_string_parts() {
    let out = run_prints(
        r#"
        fun main() {
            val text = "a1b2c3"
            println(text.count { it.isDigit() })
            println(text.count { it.isLetter() })
            println(text.count { it.isWhitespace() })
        }
    "#,
    );
    assert_eq!(out, &["3", "3", "0"]);
}

#[test]
fn test_character_slice_from_indexes() {
    let out = run_prints(
        r#"
        fun main() {
            val word = "abcdef"
            println(word[0])
            println(word[3])
            println(word[5])
            println(word[word.lastIndex])
        }
    "#,
    );
    assert_eq!(out, &["a", "d", "f", "f"]);
}

#[test]
fn test_character_try_catch_invalid_index_is_runtime_error() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "ok"
            try {
                println(value[9])
            } catch (e: Exception) {
                println("out-of-range")
            }
        }
    "#,
    );
    assert_eq!(out, &["out-of-range"]);
}

#[test]
fn test_character_codepoint_conversions() {
    let out = run_prints(
        r#"
        fun main() {
            val a = 'a'
            val b = 'b'
            println(a.code + 1)
            println((b.code - a.code))
            println(97.toChar())
            println(('A'.code - 'a'.code))
        }
    "#,
    );
    assert_eq!(out, &["98", "1", "a", "-32"]);
}

#[test]
fn test_character_join_to_string_with_transform() {
    let out = run_prints(
        r#"
        fun main() {
            val values = charArrayOf('a', 'b', 'c')
            val transformed = values.joinToString(",") { it.uppercaseChar().toString() }
            val mapped = values.map { it.uppercaseChar() }.joinToString("")
            println(transformed)
            println(mapped)
        }
    "#,
    );
    assert_eq!(out, &["A,B,C", "ABC"]);
}

#[test]
fn test_character_counting_in_filter_expressions() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "A1 b2C-3"
            println(value.count { it.isUpperCase() })
            println(value.count { it.isLowerCase() })
            println(value.count { it.isDigit() })
            println(value.count { it.isWhitespace() })
            println(value.count { it.isLetter() })
            println(value.count { it.isLetterOrDigit() })
        }
    "#,
    );
    assert_eq!(out, &["2", "1", "3", "1", "3", "6"]);
}

#[test]
fn test_character_comparison_to_string_length() {
    let out = run_prints(
        r#"
        fun main() {
            val left = "ab"
            val right = "aB"
            println(left[0] < right[1])
            println(left[1] == 'b')
            println(right.compareTo("aa"))
            println(left.compareTo(right))
        }
    "#,
    );
    assert_eq!(out, &["false", "true", "1", "-1"]);
}

#[test]
fn test_character_unicode_edge_cases() {
    let out = run_prints(
        r#"
        fun main() {
            val space = '\u0020'
            val tab = '\u0009'
            val euro = '€'
            println(space.isWhitespace())
            println(tab.isWhitespace())
            println(euro.isLetterOrDigit())
            println(euro.isDefined())
            println(euro.code)
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "false", "true", "8364"]);
}

#[test]
fn test_character_building_and_mutable_lists() {
    let out = run_prints(
        r#"
        fun main() {
            val chars = mutableListOf<Char>()
            chars.add('k')
            chars.add('o')
            chars.add('t')
            chars[1] = 'O'
            println(chars.joinToString(""))
            println(chars.size)
            println(chars[2])
        }
    "#,
    );
    assert_eq!(out, &["kOt", "3", "t"]);
}

#[test]
fn test_character_conditional_branching_on_case() {
    let out = run_prints(
        r#"
        fun main() {
            val c = 'G'
            val bucket = when {
                c.isUpperCase() -> "upper"
                c.isDigit() -> "digit"
                else -> "other"
            }
            val c2 = '4'
            val bucket2 = when {
                c2.isUpperCase() -> "upper"
                c2.isLowerCase() -> "lower"
                c2.isDigit() -> "digit"
                else -> "other"
            }
            println(bucket)
            println(bucket2)
        }
    "#,
    );
    assert_eq!(out, &["upper", "digit"]);
}
