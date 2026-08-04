use crate::helpers::run_prints;

#[test]
fn test_string_length_and_content() {
    let out = run_prints(
        r#"
        fun main() {
            val empty = ""
            val word = "Kotlin"
            println(empty.length)
            println(word.length)
            println(empty == "")
            println(word == "Kotlin")
        }
    "#,
    );
    assert_eq!(out, &["0", "6", "true", "true"]);
}

#[test]
fn test_string_mutable_append_and_reassign() {
    let out = run_prints(
        r#"
        fun main() {
            var value = "Hello"
            value += ", "
            value += "World"
            println(value)
            println(value.length)
        }
    "#,
    );
    assert_eq!(out, &["Hello, World", "12"]);
}

#[test]
fn test_string_template_expression() {
    let out = run_prints(
        r#"
        fun main() {
            val a = 3
            val b = 4
            println("$a + $b = ${a + b}")
        }
    "#,
    );
    assert_eq!(out, &["3 + 4 = 7"]);
}

#[test]
fn test_string_escape_sequence() {
    let out = run_prints(
        r#"
        fun main() {
            val quoted = "He said \"Kotlin\""
            val path = "C:\\temp\\out"
            println(quoted)
            println(path)
        }
    "#,
    );
    assert_eq!(out, &["He said \"Kotlin\"", "C:\\temp\\out"]);
}

#[test]
fn test_string_newline_escape_output() {
    let out = run_prints(
        r#"
        fun main() {
            println("a\nb")
        }
    "#,
    );
    assert_eq!(out, &["a", "b"]);
}

#[test]
fn test_string_index_and_iteration() {
    let out = run_prints(
        r#"
        fun main() {
            val word = "chat"
            var letters = ""
            for (ch in word) {
                letters += ch
            }
            println(word[0])
            println(word[3])
            println(letters)
        }
    "#,
    );
    assert_eq!(out, &["c", "t", "chat"]);
}

#[test]
fn test_string_comparison() {
    let out = run_prints(
        r#"
        fun main() {
            println("ab" < "ac")
            println("ab" == "ab")
            println("ab" != "BA")
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_string_numeric_concatenation() {
    let out = run_prints(
        r#"
        fun main() {
            val count = 2
            println("count=" + count)
            println("next=${count + 1}")
        }
    "#,
    );
    assert_eq!(out, &["count=2", "next=3"]);
}

#[test]
fn test_string_substring_and_suffix() {
    let out = run_prints(
        r#"
        fun main() {
            val source = "abcdef"
            println(source.substring(1, 4))
            println(source.substring(3))
        }
    "#,
    );
    assert_eq!(out, &["bcd", "def"]);
}

#[test]
fn test_string_trim_and_contains_like_checks() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "  Kotlin  "
            val trimmed = value.trim()
            println(trimmed)
            println(trimmed.startsWith("Kot"))
            println(trimmed.endsWith("lin"))
            println(trimmed.contains("tin"))
        }
    "#,
    );
    assert_eq!(out, &["Kotlin", "true", "true", "true"]);
}

#[test]
fn test_empty_and_blank_predicates() {
    let out = run_prints(
        r#"
        fun main() {
            val empty = ""
            val blanks = "  \n\t"
            val word = "k"
            println(empty.isEmpty())
            println(empty.isBlank())
            println(blanks.isEmpty())
            println(blanks.isBlank())
            println(word.isBlank())
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "false", "true", "false"]);
}

#[test]
fn test_nullable_string_is_null_or_empty() {
    let out = run_prints(
        r#"
        fun main() {
            val missing: String? = null
            val empty: String? = ""
            println(missing.isNullOrEmpty())
            println(empty.isNullOrEmpty())
            println(("abc").isNullOrEmpty())
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "false"]);
}

#[test]
fn test_case_transformations() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "Kotlin"
            println(value.lowercase())
            println(value.uppercase())
        }
    "#,
    );
    assert_eq!(out, &["kotlin", "KOTLIN"]);
}

#[test]
fn test_trim_start_and_end_variants() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "  Kotlin  "
            println(value.trimStart())
            println(value.trimEnd())
            println(value.trim())
        }
    "#,
    );
    assert_eq!(out, &["Kotlin  ", "  Kotlin", "Kotlin"]);
}

#[test]
fn test_string_padding_and_width() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "7"
            println(value.padStart(3, "0"))
            println(value.padEnd(4, "_"))
            println("abc".padStart(5, "x"))
        }
    "#,
    );
    assert_eq!(out, &["007", "7___", "xxabc"]);
}

#[test]
fn test_index_of_and_last_index_of() {
    let out = run_prints(
        r#"
        fun main() {
            val word = "banana"
            println(word.indexOf("na"))
            println(word.lastIndexOf("na"))
            println(word.indexOf("na", 3))
            println(word.indexOf("x"))
        }
    "#,
    );
    assert_eq!(out, &["2", "4", "4", "-1"]);
}

#[test]
fn test_replace_and_replace_first() {
    let out = run_prints(
        r#"
        fun main() {
            val word = "banana"
            println(word.replace("na", "NA"))
            println(word.replaceFirst("ba", "BO"))
            println(word.replace("na", "", false))
        }
    "#,
    );
    assert_eq!(out, &["baNANA", "NAna", "baa"]);
}

#[test]
fn test_split_and_reconstruct() {
    let out = run_prints(
        r#"
        fun main() {
            val parts = "a,b,c,d".split(",")
            println(parts.size)
            println(parts[0])
            println(parts[3])
            println(parts.joinToString("|"))
        }
    "#,
    );
    assert_eq!(out, &["4", "a", "d", "a|b|c|d"]);
}

#[test]
fn test_reverse_and_reversed_roundtrip() {
    let out = run_prints(
        r#"
        fun main() {
            val word = "abcd"
            val reversed = word.reversed()
            println(reversed)
            println(reversed.reversed())
        }
    "#,
    );
    assert_eq!(out, &["dcba", "abcd"]);
}

#[test]
fn test_contains_with_ignore_case() {
    let out = run_prints(
        r#"
        fun main() {
            val word = "Kotlin"
            println(word.contains("kot"))
            println(word.contains("kot", true))
            println(word.startsWith("Ko"))
            println(word.endsWith("IN", ignoreCase = true))
        }
    "#,
    );
    assert_eq!(out, &["false", "true", "true", "true"]);
}

#[test]
fn test_repeat_and_slice() {
    let out = run_prints(
        r#"
        fun main() {
            println("ha".repeat(3))
            println("kotlin".slice(1..3))
            println("kotlin".slice(IntRange(0, 2)))
        }
    "#,
    );
    assert_eq!(out, &["hahaha", "otl", "kot"]);
}

#[test]
fn test_take_drop_prefix_suffix() {
    let out = run_prints(
        r#"
        fun main() {
            val word = "abcdef"
            println(word.take(3))
            println(word.drop(3))
            println(word.takeLast(2))
            println(word.dropLast(4))
        }
    "#,
    );
    assert_eq!(out, &["abc", "def", "ef", "ab"]);
}

#[test]
fn test_last_index_property_and_character_access() {
    let out = run_prints(
        r#"
        fun main() {
            val word = "rust"
            println(word.lastIndex)
            println(word[0])
            println(word[word.lastIndex])
        }
    "#,
    );
    assert_eq!(out, &["3", "r", "t"]);
}

#[test]
fn test_substring_bounds_and_stepwise_offsets() {
    let out = run_prints(
        r#"
        fun main() {
            val word = "compiler"
            println(word.substring(0, 3))
            println(word.substring(3))
            println(word.substring(word.length - 2))
        }
    "#,
    );
    assert_eq!(out, &["com", "iler", "er"]);
}

#[test]
fn test_string_template_braces_complex_expression() {
    let out = run_prints(
        r#"
        fun main() {
            val width = 4
            val height = 5
            println("${width}x${height}=${width * height}")
            println("$${width + height}")
        }
    "#,
    );
    assert_eq!(out, &["4x5=20", "$9"]);
}

#[test]
fn test_string_interpolation_with_null_fallback() {
    let out = run_prints(
        r#"
        fun main() {
            val nullable: String? = null
            val present: String? = "value"
            println("[$nullable]")
            println("[${nullable ?: "fallback"}]")
            println("[${present ?: "fallback"}]")
        }
    "#,
    );
    assert_eq!(out, &["[null]", "[fallback]", "[value]"]);
}

#[test]
fn test_string_comparison_and_ordering() {
    let out = run_prints(
        r#"
        fun main() {
            println("abc" < "abd")
            println("abc" == "abc")
            println("ABC" < "abc")
            println("ABC".equals("abc", ignoreCase = true))
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true", "true"]);
}

#[test]
fn test_string_builder_mutability() {
    let out = run_prints(
        r#"
        fun main() {
            val builder = StringBuilder()
            builder.append("a")
            builder.append("-")
            builder.append("z")
            builder.insert(2, "middle")
            println(builder.toString())
            println(builder.length)
        }
    "#,
    );
    assert_eq!(out, &["a-middlez", "9"]);
}

#[test]
fn test_compare_to_numeric_string_lengths() {
    let out = run_prints(
        r#"
        fun main() {
            val words = listOf("a", "ab", "abc")
            var shortest = words[0]
            for (word in words) {
                if (word.length < shortest.length) {
                    shortest = word
                }
            }
            println(shortest)
        }
    "#,
    );
    assert_eq!(out, &["a"]);
}

#[test]
fn test_string_filter_digits_and_letters() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "a1b2c3"
            val digits = value.filter { it.isDigit() }
            val letters = value.filter { it.isLetter() }
            println(digits)
            println(letters)
        }
    "#,
    );
    assert_eq!(out, &["123", "abc"]);
}

#[test]
fn test_compare_to_on_empty_and_singleton() {
    let out = run_prints(
        r#"
        fun main() {
            println("".isEmpty())
            println("".isNotEmpty())
            println("".compareTo(""))
            println("a".compareTo(""))
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "0", "1"]);
}

#[test]
fn test_raw_string_multiline_and_indentation_removal() {
    let out = run_prints(
        r#"
        fun main() {
            val value = """
                line one
                line two
            """.trimIndent()
            println(value)
            println(value.lines().size)
        }
    "#,
    );
    assert_eq!(out, &["line one", "line two", "2"]);
}

#[test]
fn test_trim_margin_removes_custom_delimiter() {
    let out = run_prints(
        r#"
        fun main() {
            val value = """
                |k
                |otlin
                """.trimMargin("|")
            println(value)
            println(value.lines().size)
        }
    "#,
    );
    assert_eq!(out, &["k", "otlin", "2"]);
}

#[test]
fn test_substring_invalid_range_throws() {
    let out = run_prints(
        r#"
        fun main() {
            val word = "abc"
            try {
                println(word.substring(5))
            } catch (e: Exception) {
                println("error")
            }
        }
    "#,
    );
    assert_eq!(out, &["error"]);
}

#[test]
fn test_remove_prefix_and_suffix_are_idempotent_when_absent() {
    let out = run_prints(
        r#"
        fun main() {
            val word = "kotlin"
            println(word.removePrefix("ko"))
            println(word.removePrefix("x"))
            println(word.removeSuffix("in"))
            println(word.removeSuffix("x"))
        }
    "#,
    );
    assert_eq!(out, &["tlin", "kotlin", "kotl", "kotlin"]);
}

#[test]
fn test_split_with_limit_and_regex_whitespace() {
    let out = run_prints(
        r#"
        fun main() {
            val parts = "a b  c d".trim().split("\\s+".toRegex(), 3)
            println(parts.size)
            println(parts[0])
            println(parts[1])
            println(parts[2])
        }
    "#,
    );
    assert_eq!(out, &["3", "a", "b", "c d"]);
}

#[test]
fn test_string_to_byte_array_roundtrip() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = "ok".toByteArray()
            val back = String(bytes)
            println(bytes.size)
            println(back)
        }
    "#,
    );
    assert_eq!(out, &["2", "ok"]);
}

#[test]
fn test_string_repeat_when_count_zero() {
    let out = run_prints(
        r#"
        fun main() {
            println("x".repeat(0))
            println("x".padStart(0))
            println("x".padEnd(0))
        }
    "#,
    );
    assert_eq!(out, &["", "x", "x"]);
}

#[test]
fn test_substring_before_and_after_helpers() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "name=value"
            println(value.substringAfter("="))
            println(value.substringBefore("="))
            println("plain".substringAfter("=", "missing"))
        }
    "#,
    );
    assert_eq!(out, &["value", "name", "missing"]);
}

#[test]
fn test_substring_before_last_and_after_last_boundaries() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "a/b/c"
            println(value.substringBeforeLast("/"))
            println(value.substringAfterLast("/"))
            println("nodelim".substringAfterLast("/", "none"))
            println("x/y/".substringAfterLast("/"))
        }
    "#,
    );
    assert_eq!(out, &["a/b", "c", "none", ""]);
}

#[test]
fn test_replace_range_and_region_matches() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "abcdef"
            println(value.replaceRange(1, 3, "ZZ"))
            println(value.regionMatches(1, "CD", 0, 2, ignoreCase = true))
            println(value.regionMatches(1, "Cd", 0, 2, ignoreCase = true))
        }
    "#,
    );
    assert_eq!(out, &["aZZdef", "true", "false"]);
}

#[test]
fn test_lines_and_trim_with_blank_lines() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "a\n\nb\n"
            val raw = value.lines()
            println(raw.size)
            println(raw[1])
            println(raw[2].isEmpty())
            println(value.lines().filter { it.isNotEmpty() }.joinToString("|"))
        }
    "#,
    );
    assert_eq!(out, &["3", "", "true", "a|b"]);
}

#[test]
fn test_to_list_of_chars_roundtrip() {
    let out = run_prints(
        r#"
        fun main() {
            val chars = "dog".toCharArray()
            println(chars.joinToString(","))
            var rebuilt = ""
            for (ch in chars) {
                rebuilt += ch
            }
            println(rebuilt)
        }
    "#,
    );
    assert_eq!(out, &["d,o,g", "dog"]);
}

#[test]
fn test_string_prefix_and_suffix_navigation() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "api/v1/resource"
            println(value.substringBefore("/"))
            println(value.substringAfter("/"))
            println(value.substringAfterLast("/"))
            println(value.substringBeforeLast("/"))
            println(value.substringBefore("x", "missing"))
            println(value.substringAfter("x", "missing"))
        }
    "#,
    );
    assert_eq!(
        out,
        &[
            "api",
            "v1/resource",
            "resource",
            "api/v1",
            "missing",
            "missing"
        ]
    );
}

#[test]
fn test_string_replace_range_and_region() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "kotlin"
            println(value.replaceRange(1, 4, "A"))
            println(value.replaceRange(0, 1, "Z"))
            println(value.replaceFirst("li", "LI"))
        }
    "#,
    );
    assert_eq!(out, &["kAin", "Zotlin", "kotLIn"]);
}

#[test]
fn test_string_chunked_and_windowed_views() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "abcdef"
            println(value.chunked(2).joinToString("|"))
            val windows = value.windowed(3, 2)
            println(windows.joinToString("|"))
        }
    "#,
    );
    assert_eq!(out, &["ab|cd|ef", "abc|cde"]);
}

#[test]
fn test_string_take_while_and_drop_while_boundaries() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "12abc34"
            println(value.takeWhile { it.isDigit() })
            println(value.dropWhile { it.isDigit() })
            println(value.takeLastWhile { it.isDigit() })
            println(value.dropLastWhile { it.isDigit() })
        }
    "#,
    );
    assert_eq!(out, &["12", "abc34", "34", "12abc"]);
}

#[test]
fn test_string_char_at_invalid_index_throws() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "kotlin"
            try {
                println(value[10])
            } catch (e: java.lang.StringIndexOutOfBoundsException) {
                println("out_of_bounds")
            }
        }
    "#,
    );
    assert_eq!(out, &["out_of_bounds"]);
}

#[test]
fn test_string_line_splitting_retains_trailing_empty() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "a\nb\n"
            val lines = value.lines()
            println(lines.size)
            println(lines[0])
            println(lines[1])
            println(lines[2])
        }
    "#,
    );
    assert_eq!(out, &["3", "a", "b", ""]);
}

#[test]
fn test_string_padstart_and_padend_short_circuit_when_width_too_small() {
    let out = run_prints(
        r#"
        fun main() {
            println("abcdef".padStart(3, "x"))
            println("abcdef".padEnd(2, "x"))
            println("a".padStart(3, "."))
            println("a".padEnd(3, "."))
        }
    "#,
    );
    assert_eq!(out, &["abcdef", "abcdef", "..a", "a.."]);
}

#[test]
fn test_string_repeat_negative_throws() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                println("x".repeat(-1))
            } catch (e: Exception) {
                println("repeat-error")
            }
        }
    "#,
    );
    assert_eq!(out, &["repeat-error"]);
}

#[test]
fn test_string_slice_out_of_bounds_throws() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "abc"
            try {
                println(value.slice(2..5))
            } catch (e: Exception) {
                println("slice-error")
            }
        }
    "#,
    );
    assert_eq!(out, &["slice-error"]);
}

#[test]
fn test_string_filter_and_counted_predicates() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "a1b2c3d4"
            val digits = value.count { it.isDigit() }
            val letters = value.count { it.isLetter() }
            val filtered = value.filterIndexed { index, ch -> index % 2 == 0 && ch.isLetter() }
            println(digits)
            println(letters)
            println(filtered)
        }
    "#,
    );
    assert_eq!(out, &["4", "4", "ac"]);
}
