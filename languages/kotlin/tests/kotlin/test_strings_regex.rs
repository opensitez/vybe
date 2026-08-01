use crate::helpers::run_prints;

#[test]
fn test_regex_match_and_non_match() {
    let out = run_prints(
        r#"
        fun main() {
            val number = Regex("\\d+")
            println(number.matches("12345"))
            println(number.matches("12a45"))
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_string_to_regex_matches() {
    let out = run_prints(
        r#"
        fun main() {
            val number = "\\d{2,3}".toRegex()
            println(number.matches("42"))
            println(number.matches("4"))
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_regex_match_entire_boundary() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("[a-z]+")
            println(pattern.matchEntire("abc") != null)
            println(pattern.matchEntire("abc123") != null)
            println(pattern.matches("abc"))
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "true"]);
}

#[test]
fn test_regex_find_first_match() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("\\d+")
            val result = pattern.find("id-42-code")
            println(result?.value ?: "none")
            println(result?.range?.first ?: -1)
            println(result?.range?.last ?: -1)
        }
    "#,
    );
    assert_eq!(out, &["42", "3", "4"]);
}

#[test]
fn test_regex_find_all_numbers() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("\\d+")
            val values = pattern.findAll("a1b22c333")
            println(values.count())
            println(values.joinToString("|") { it.value })
        }
    "#,
    );
    assert_eq!(out, &["3", "1|22|333"]);
}

#[test]
fn test_regex_find_on_starting_index() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("\\d+")
            val first = pattern.find("a1b22", 2)
            println(first?.value ?: "none")
        }
    "#,
    );
    assert_eq!(out, &["22"]);
}

#[test]
fn test_regex_matches_at_offset() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("\\d+")
            println(pattern.matchesAt("abc123", 3))
            println(pattern.matchesAt("abc123", 0))
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_regex_contains_match() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("cat|dog")
            println(pattern.containsMatchIn("the catalog"))
            println(pattern.containsMatchIn("fish"))
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_regex_replace_literal_and_first() {
    let out = run_prints(
        r#"
        fun main() {
            val value = "a+b+c"
            val escaped = Regex.escape(value)
            println(escaped)
            println(Regex(escaped).replace(value, "_"))
            println(Regex("\\d+").replaceFirst("x1x2x", "NUM"))
        }
    "#,
    );
    assert_eq!(out, &["\\Qa+b+c\\E", "_", "xNUMx2x"]);
}

#[test]
fn test_regex_replace_with_transform() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("(\\d+)")
            val out = pattern.replace("a12b34c") { match -> match.value.reversed() }
            println(out)
        }
    "#,
    );
    assert_eq!(out, &["a21b43c"]);
}

#[test]
fn test_regex_replace_with_group_mapping() {
    let out = run_prints(
        r#"
        fun main() {
            val input = "id:42:done"
            val pattern = Regex("id:(\\d+):done")
            val output = pattern.replace(input) { match ->
                "ID=" + match.groups[1]!!.value
            }
            println(output)
        }
    "#,
    );
    assert_eq!(out, &["ID=42"]);
}

#[test]
fn test_regex_split_simple() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("\\s+")
            val parts = pattern.split("a  b   c")
            println(parts.size)
            println(parts.joinToString("|"))
        }
    "#,
    );
    assert_eq!(out, &["3", "a|b|c"]);
}

#[test]
fn test_regex_split_with_limit() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex(",")
            val parts = pattern.split("a,b,c,d", limit = 2)
            println(parts.size)
            println(parts[0])
            println(parts[1])
        }
    "#,
    );
    assert_eq!(out, &["2", "a", "b,c,d"]);
}

#[test]
fn test_regex_split_to_sequence() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("-")
            val parts = pattern.splitToSequence("x-y-z").toList()
            println(parts.size)
            println(parts.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["3", "x,y,z"]);
}

#[test]
fn test_regex_option_ignore_case() {
    let out = run_prints(
        r#"
        fun main() {
            val lower = Regex("cat", RegexOption.IGNORE_CASE)
            println(lower.matches("CAT"))
            println(lower.matches("dog"))
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_regex_option_multiline() {
    let out = run_prints(
        r#"
        fun main() {
            val anchored = Regex("^kotlin$", RegexOption.MULTILINE)
            println(anchored.containsMatchIn("java\nkotlin\nrust"))
            println(anchored.containsMatchIn("kotlin "))
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_regex_option_dot_matches_all() {
    let out = run_prints(
        r#"
        fun main() {
            val dot = Regex("a.*c")
            val dotAny = Regex("a.*c", RegexOption.DOT_MATCHES_ALL)
            val text = "a\nc"
            println(dot.matches(text))
            println(dotAny.matches(text))
        }
    "#,
    );
    assert_eq!(out, &["false", "true"]);
}

#[test]
fn test_regex_option_set_builder() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("k.t", setOf(RegexOption.IGNORE_CASE, RegexOption.DOT_MATCHES_ALL))
            println(pattern.matches("K\nt"))
        }
    "#,
    );
    assert_eq!(out, &["true"]);
}

#[test]
fn test_regex_capture_groups_positional() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("(\\d{2})-(\\w+)")
            val result = pattern.find("42-abc")
            println(result?.destructured?.component1() ?: "none")
            println(result?.destructured?.component2() ?: "none")
            println(result?.groupValues?.size ?: 0)
        }
    "#,
    );
    assert_eq!(out, &["42", "abc", "3"]);
}

#[test]
fn test_regex_optional_capture_groups() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("(\\d+)?-(\\w*)")
            val result = pattern.find(" -abc")
            println(result?.groupValues?.get(1) ?: "missing")
            println(result?.groupValues?.get(2) ?: "missing")
        }
    "#,
    );
    assert_eq!(out, &["", "abc"]);
}

#[test]
fn test_regex_start_end_indices() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("dog")
            val result = pattern.find("The dog runs")
            println(result?.range?.start ?: -1)
            println(result?.range?.endInclusive ?: -1)
        }
    "#,
    );
    assert_eq!(out, &["4", "6"]);
}

#[test]
fn test_regex_find_all_with_indexes() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("[A-Za-z]")
            val first = pattern.find("A1b2C3")
            println(first?.value ?: "none")
            val all = pattern.findAll("A1b2C3").toList()
            println(all[0].range.start)
            println(all[1].range.start)
            println(all[2].range.start)
        }
    "#,
    );
    assert_eq!(out, &["A", "0", "2", "4"]);
}

#[test]
fn test_regex_from_literal_treats_as_plain_text() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex.fromLiteral("a+b")
            println(pattern.containsMatchIn("c a+b d"))
            println(pattern.matches("a+b"))
            println(pattern.matches("aaab"))
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "false"]);
}

#[test]
fn test_regex_replace_with_counted_callback() {
    let out = run_prints(
        r#"
        fun main() {
            var index = 0
            val pattern = Regex("\\d")
            val output = pattern.replace("a1b2c3") { match ->
                val value = "${index}:${match.value}"
                index += 1
                value
            }
            println(output)
            println(index)
        }
    "#,
    );
    assert_eq!(out, &["a0:1b1:2c2:3", "3"]);
}

#[test]
fn test_regex_find_all_distinct_via_set() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("\\b\\w+")
            val words = pattern.findAll("a a b c a").map { it.value }.toList().toSet()
            println(words.size)
            println(words.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["4", "a,b,c"]);
}

#[test]
fn test_regex_grouping_and_replacement_with_backreference() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("(\\w+)\\s+\\1")
            println(pattern.matches("go go"))
            println(pattern.matches("go now"))
            val text = pattern.replace("yo yo test") { match ->
                "[${match.groupValues[1]}]"
            }
            println(text)
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "[yo] test"]);
}

#[test]
fn test_regex_multiple_captures_mapping() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("(\\d)(\\w)(\\w)")
            val result = pattern.find("9ab")
            val mapped = result?.groups?.let {
                "${it[1]?.value}-${it[2]?.value}-${it[3]?.value}"
            } ?: "none"
            println(mapped)
            println(pattern.find("abc") == null)
        }
    "#,
    );
    assert_eq!(out, &["9-a-b", "true"]);
}

#[test]
fn test_regex_empty_match_handling() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("a*")
            val result = pattern.matchEntire("")
            println(result != null)
            println(result?.value ?: "none")
        }
    "#,
    );
    assert_eq!(out, &["true", ""]);
}

#[test]
fn test_regex_invalid_pattern_error() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                val pattern = Regex("[")
                println(pattern)
            } catch (e: java.lang.RuntimeException) {
                println("bad")
            }
        }
    "#,
    );
    assert_eq!(out, &["bad"]);
}

#[test]
fn test_regex_matcher_with_all_occurrences_and_matcher_state() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("\\b\\w+\\b")
            val matcher = pattern.toPattern().matcher("one two three")
            var matches = ""
            while (matcher.find()) {
                matches += matcher.group()
                matches += ":"
                matches += matcher.start().toString()
                matches += "-"
                matches += matcher.end().toString()
                matches += ";"
            }
            println(matches)
        }
    "#,
    );
    assert_eq!(out, &["one:0-3;two:4-7;three:8-13;"]);
}

#[test]
fn test_regex_find_returns_null_when_absent() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("x+")
            val result = pattern.find("abc")
            println(result == null)
            val value = pattern.find("abc")?.value ?: "missing"
            println(value)
        }
    "#,
    );
    assert_eq!(out, &["true", "missing"]);
}

#[test]
fn test_regex_replace_first_only() {
    let out = run_prints(
        r##"
        fun main() {
            val pattern = Regex("\\d+")
            println(pattern.replaceFirst("a1b22c333", "#"))
            println(pattern.replaceFirst("abc", "#"))
        }
    "##,
    );
    assert_eq!(out, &["a#b22c333", "abc"]);
}

#[test]
fn test_regex_split_retains_trailing_empties_without_limit() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex(",")
            val parts = pattern.split("a,b,c,")
            println(parts.size)
            println(parts.joinToString("|"))
        }
    "#,
    );
    assert_eq!(out, &["4", "a|b|c|"]);
}

#[test]
fn test_regex_split_with_unicode_set_and_limit_zero() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("\\s+")
            val parts = pattern.split("a b  c", limit = 0)
            println(parts.size)
            println(parts.joinToString("|"))
        }
    "#,
    );
    assert_eq!(out, &["3", "a|b|c"]);
}

#[test]
fn test_regex_named_capture_group_access() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("(?<id>\\d+)-(?<name>\\w+)")
            val result = pattern.find("42-kotlin")
            println(result?.groups?.get("id")?.value ?: "missing")
            println(result?.groups?.get("name")?.value ?: "missing")
        }
    "#,
    );
    assert_eq!(out, &["42", "kotlin"]);
}

#[test]
fn test_regex_unicode_class_and_case_option_interaction() {
    let out = run_prints(
        r#"
        fun main() {
            val letters = Regex("straße", RegexOption.IGNORE_CASE)
            println(letters.matches("STRASSE"))
            println(letters.matchesAt("XXstraßeYY", 2))
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_regex_literal_option_treats_regex_meta_as_text() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("a+b|c*", RegexOption.LITERAL)
            println(pattern.matches("a+b|c*"))
            println(pattern.matches("aaab"))
            println(pattern.containsMatchIn("xxa+b|c*yy"))
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "true"]);
}

#[test]
fn test_regex_find_all_uses_capturing_groups() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("(\\w+)(\\d)")
            val matches = pattern.findAll("a1 b22 c3")
            var output = ""
            for (item in matches) {
                output += item.groupValues[1]
                output += "-"
                output += item.groupValues[2]
                output += ";"
            }
            println(output)
        }
    "#,
    );
    assert_eq!(out, &["a-1;b2-2;c-3;"]);
}

#[test]
fn test_regex_find_with_start_index_and_no_match() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("\\d+")
            println(pattern.find("abc123", 1)?.value ?: "none")
            println(pattern.find("abc123", 4) == null)
            println(pattern.find("abc", 3) == null)
            println(pattern.find("abc", 5) == null)
        }
    "#,
    );
    assert_eq!(out, &["123", "true", "true", "true"]);
}

#[test]
fn test_regex_to_pattern_with_java_matcher_groups_and_positions() {
    let out = run_prints(
        r#"
        fun main() {
            val matcher = Regex("(\\w)(\\d)").toPattern().matcher("a1 b2 c3")
            var trace = ""
            while (matcher.find()) {
                trace += matcher.group(1)
                trace += matcher.group(2)
                trace += matcher.start().toString()
                trace += matcher.end().toString()
                trace += "|"
            }
            println(trace)
        }
    "#,
    );
    assert_eq!(out, &["a102|b235|c367|"]);
}

#[test]
fn test_regex_split_keeps_trailing_empty_when_limit_negative() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex(",")
            val parts = pattern.split("a,b,c,", limit = -1)
            println(parts.size)
            println(parts.joinToString("|"))
        }
    "#,
    );
    assert_eq!(out, &["4", "a|b|c|"]);
}

#[test]
fn test_regex_replace_with_group_reference_tokens() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("(\\w)(\\d)")
            println(pattern.replace("a1 b2", "\$2-\$1"))
        }
    "#,
    );
    assert_eq!(out, &["1-a 2-b"]);
}

#[test]
fn test_regex_find_all_on_empty_input() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("\\d+")
            val values = pattern.findAll("").toList()
            println(values.size)
        }
    "#,
    );
    assert_eq!(out, &["0"]);
}

#[test]
fn test_regex_matcher_reset_reuse_works() {
    let out = run_prints(
        r#"
        fun main() {
            val matcher = Regex("a(\\d)").toPattern().matcher("a1x")
            println(matcher.find())
            println(matcher.group(1))

            matcher.reset("a2")
            println(matcher.find())
            println(matcher.group(1))
        }
    "#,
    );
    assert_eq!(out, &["true", "1", "true", "2"]);
}

#[test]
fn test_regex_find_all_distinguishes_full_match_and_groups() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex("(\\w)-(\\d)")
            val first = pattern.find("a-1")
            println(first?.groupValues?.getOrNull(0) ?: "none")
            println(first?.groupValues?.getOrNull(1) ?: "none")
            println(first?.groupValues?.getOrNull(2) ?: "none")
            println(first?.groupValues?.size ?: 0)
        }
    "#,
    );
    assert_eq!(out, &["a-1", "a", "1", "3"]);
}

#[test]
fn test_regex_split_limit_one_behavior() {
    let out = run_prints(
        r#"
        fun main() {
            val pattern = Regex(",")
            val parts = pattern.split("a,b,c", 1)
            println(parts.size)
            println(parts[0])
        }
    "#,
    );
    assert_eq!(out, &["1", "a,b,c"]);
}
