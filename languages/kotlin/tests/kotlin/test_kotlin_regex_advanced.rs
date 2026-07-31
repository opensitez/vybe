use crate::helpers::run_prints;

#[test]
fn test_regex_ignore_case_option() {
    let out = run_prints(r#"
        fun main() {
            val re = "ab+c".toRegex(RegexOption.IGNORE_CASE)
            println(re.matches("ABBC"))
            println(re.matches("abc"))
            println(re.matches("ABC"))
        }
    "#);
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_regex_multiline_anchors() {
    let out = run_prints(r#"
        fun main() {
            val text = "x\ny\nz"
            val withAnchors = "^y$".toRegex(RegexOption.MULTILINE)
            val matched = withAnchors.containsMatchIn(text)
            val bad = "^y$".toRegex().containsMatchIn(text)
            println(matched)
            println(bad)
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_regex_dot_matches_all() {
    let out = run_prints(r#"
        fun main() {
            val text = "a\nb"
            val normal = Regex("a.b")
            val dotAll = Regex("a.b", RegexOption.DOT_MATCHES_ALL)
            println(normal.containsMatchIn(text))
            println(dotAll.containsMatchIn(text))
        }
    "#);
    assert_eq!(out, &["false", "true"]);
}

#[test]
fn test_regex_comments_and_literal_mode() {
    let out = run_prints(r#"
        fun main() {
            val token = "a#b".toRegex(setOf(RegexOption.COMMENTS))
            val literal = Regex("a#b", RegexOption.LITERAL)
            println(token.matches("a#b"))
            println(literal.matches("a#b"))
            println(token.matches("a b"))
        }
    "#);
    assert_eq!(out, &["false", "true", "false"]);
}

#[test]
fn test_regex_capture_groups_and_names() {
    let out = run_prints(r#"
        fun main() {
            val regex = Regex("^(?<name>[a-z]+):(?<value>\\d+)$")
            val result = regex.find("age:42")
            println(result != null)
            val groups = result?.groups
            println(groups?.size)
            println(groups?.get("name")?.value)
            println(groups?.get("value")?.value)
        }
    "#);
    assert_eq!(out, &["true", "3", "age", "42"]);
}

#[test]
fn test_regex_find_all_values() {
    let out = run_prints(r#"
        fun main() {
            val regex = Regex("\\d+")
            val found = regex.findAll("a1 b22 c333")
            println(found.map { it.value }.joinToString(","))
            println(found.count())
        }
    "#);
    assert_eq!(out, &["1,22,333", "3"]);
}

#[test]
fn test_regex_replace_with_transform() {
    let out = run_prints(r#"
        fun main() {
            val regex = Regex("(\\w)(\\d+)")
            val replaced = regex.replace("a1 b22", { match ->
                match.groupValues[1] + "=" + match.groupValues[2]
            })
            println(replaced)
        }
    "#);
    assert_eq!(out, &["a=1 b=22"]);
}

#[test]
fn test_regex_split_keep_delimiters() {
    let out = run_prints(r#"
        fun main() {
            val regex = Regex("[;,]")
            val parts = regex.split("a,b;c,d")
            println(parts.joinToString("|"))
            println(parts.size)
        }
    "#);
    assert_eq!(out, &["a|b|c|d", "4"]);
}

#[test]
fn test_regex_match_entire_line_with_options_combo() {
    let out = run_prints(r#"
        fun main() {
            val pattern = Regex("^\n*OK\\?$")
            val withComments = Regex("^\n*OK\\?$")
            println(pattern.matches("OK?"))
            println(withComments.matches("\n\nOK?"))
            val withOption = Regex("^\n*OK\\?$", RegexOption.MULTILINE)
            println(withOption.matches("line1\nOK?"))
        }
    "#);
    assert_eq!(out, &["true", "false", "false"]);
}

#[test]
fn test_regex_contains_match_in() {
    let out = run_prints(r#"
        fun main() {
            val regex = Regex("foo")
            println(regex.containsMatchIn("bar foo baz"))
            println(regex.matches("foo"))
            println(regex.matches("bar foo baz"))
        }
    "#);
    assert_eq!(out, &["true", "true", "false"]);
}

#[test]
fn test_regex_find_and_range() {
    let out = run_prints(r#"
        fun main() {
            val regex = Regex("b")
            val match = regex.find("abca", 0)
            println(match?.value)
            println(match?.range?.first)
            println(match?.range?.last)
        }
    "#);
    assert_eq!(out, &["b", "1", "1"]);
}
