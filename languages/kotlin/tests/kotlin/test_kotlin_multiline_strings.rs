kotlin_run_test!(
    test_raw_string_preserves_newlines,
    r#"
        fun main() {
            val text = """
line1
line2
line3
"""
            println(text.lines().size)
            println(text.lines()[1])
        }
    "#,
    &["4", "line2"]
);

kotlin_run_test!(
    test_trim_margin_with_custom_delimiter,
    r#"
        fun main() {
            val text = """
>one
>two
>three
""".trimMargin(">")
            println(text)
        }
    "#,
    &["one\ntwo\nthree"]
);

kotlin_run_test!(
    test_raw_string_with_quoted_marker,
    r#"
        fun main() {
            val text = """
"quoted"
not quoted
"""
            println(text.trim().split("\n")[0])
        }
    "#,
    &["\"quoted\""]
);

kotlin_run_test!(
    test_multiline_expression_embedded_interpolation,
    r#"
        fun main() {
            val n = 2
            val message = """
${'$'}n squared is ${'$'}{n * n}
${'$'}n cubed is ${'$'}{n * n * n}
"""
            val lines = message.trim().lines()
            println(lines[0])
            println(lines[1])
        }
    "#,
    &["2 squared is 4", "2 cubed is 8"]
);

kotlin_run_test!(
    test_raw_string_with_triple_quotes_escape,
    r#"
        fun main() {
            val text = """contains ${"""} inside"""
            println(text)
        }
    "#,
    &["contains  inside"]
);

kotlin_run_test!(
    test_raw_string_with_tabs_preserved,
    r#"
        fun main() {
            val text = """a\tb\tc"""
            println(text.length)
        }
    "#,
    &["5"]
);

kotlin_run_test!(
    test_raw_string_leading_spaces_trimmed,
    r#"
        fun main() {
            val text = """
            left
            right
            """.trimIndent()
            println(text)
        }
    "#,
    &["left\nright"]
);

kotlin_run_test!(
    test_raw_string_join_with_pipe,
    r#"
        fun main() {
            val lines = """a|b|c"""
            val parts = lines.split("|")
            println(parts.size)
            println(parts.joinToString(","))
        }
    "#,
    &["3", "a,b,c"]
);

kotlin_run_test!(
    test_raw_string_boolean_block,
    r#"
        fun main() {
            val ok = true
            val text = """
status=${'$'}{if (ok) "yes" else "no"}
"""
            println(text.trim())
        }
    "#,
    &["status=yes"]
);
