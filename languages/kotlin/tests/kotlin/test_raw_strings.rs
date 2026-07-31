kotlin_run_test!(
    test_raw_string_simple_multiline,
    r#"
        fun main() {
            val text = """line one
line two"""
            println(text.lines().size)
        }
    "#,
    &["2"]
);

kotlin_run_test!(
    test_raw_string_keeps_indentation,
    r#"
        fun main() {
            val text = """  a
  b"""
            println(text[0])
            println(text[3])
        }
    "#,
    &[" ", "b"]
);

kotlin_run_test!(
    test_raw_string_with_quotes,
    r#"
        fun main() {
            val text = """He said "hello""""
            println(text)
        }
    "#,
    &["He said \"hello\""]
);

kotlin_run_test!(
    test_raw_string_with_backticks,
    r#"
        fun main() {
            val text = """raw `code` value"""
            println(text)
        }
    "#,
    &["raw `code` value"]
);

kotlin_run_test!(
    test_raw_string_with_dollar_escape,
    r#"
        fun main() {
            val text = """price ${'$'}100"""
            println(text)
        }
    "#,
    &["price $100"]
);

kotlin_run_test!(
    test_raw_string_with_interpolated_expression,
    r#"
        fun main() {
            val user = "k"
            val text = """user=${user}"""
            println(text)
        }
    "#,
    &["user=k"]
);

kotlin_run_test!(
    test_raw_string_with_nested_delimiters,
    r#"
        fun main() {
            val text = """one
"""
            println(text.length)
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_raw_string_trim_indent,
    r#"
        fun main() {
            val text = """
                one
                two
                three
            """.trimIndent()
            println(text.lines().size)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_raw_string_trim_margin,
    r#"
        fun main() {
            val text = """
                |one
                |two
                |three
            """.trimMargin()
            println(text)
        }
    "#,
    &["one\ntwo\nthree"]
);

kotlin_run_test!(
    test_raw_string_join_with_plus,
    r#"
        fun main() {
            val a = """a"""
            val b = """b"""
            println(a + b)
        }
    "#,
    &["ab"]
);

kotlin_run_test!(
    test_raw_string_concat_multiple,
    r#"
        fun main() {
            val text = """a""" + """b""" + """c"""
            println(text)
        }
    "#,
    &["abc"]
);

kotlin_run_test!(
    test_raw_string_empty,
    r#"
        fun main() {
            val text = """"""
            println(text.isEmpty())
        }
    "#,
    &["true"]
);

kotlin_run_test!(
    test_raw_string_single_line_with_spaces,
    r#"
        fun main() {
            val text = """  a b  """
            println(text.trim())
        }
    "#,
    &["a b"]
);

kotlin_run_test!(
    test_raw_string_contains_newline,
    r#"
        fun main() {
            val text = """A
B
"""
            println(text.endsWith("\n"))
        }
    "#,
    &["true"]
);

kotlin_run_test!(
    test_raw_string_empty_line_middle,
    r#"
        fun main() {
            val text = """a

b"""
            println(text.lines().size)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_raw_string_char_index,
    r#"
        fun main() {
            val text = """abc"""
            println(text[1])
        }
    "#,
    &["b"]
);

kotlin_run_test!(
    test_raw_string_utf8_unicode_sequence,
    r#"
        fun main() {
            val text = """αβγ"""
            println(text.length)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_raw_string_backslash_preserved,
    r#"
        fun main() {
            val text = """x\\y"""
            println(text)
        }
    "#,
    &["x\\y"]
);

kotlin_run_test!(
    test_raw_string_with_hash_character,
    r#"
        fun main() {
            val text = """v#1"""
            println(text)
        }
    "#,
    &["v#1"]
);

kotlin_run_test!(
    test_raw_string_compare_length,
    r#"
        fun main() {
            val text = """ab
cd"""
            println(text.length)
        }
    "#,
    &["5"]
);

kotlin_run_test!(
    test_raw_string_trim_indent_complex,
    r#"
        fun main() {
            val text = """
                    a
                      b
                    c
            """.trimIndent()
            println(text.length)
        }
    "#,
    &["9"]
);

kotlin_run_test!(
    test_raw_string_with_quote_repetition,
    r#"
        fun main() {
            val open = "\"\"\""
            val text = """contains """ + open + """ inside"""
            println(text)
        }
    "#,
    &["contains \"\"\" inside"]
);

kotlin_run_test!(
    test_raw_string_with_indented_margin_custom,
    r#"
        fun main() {
            val text = """
                >a
                >b
            """.trimMargin(">")
            println(text.lines().size)
        }
    "#,
    &["2"]
);

kotlin_run_test!(
    test_raw_string_with_dollar_but_no_expr,
    r#"
        fun main() {
            val x = 1
            val text = """$${x}"""
            println(text)
        }
    "#,
    &["$1"]
);

kotlin_run_test!(
    test_raw_string_empty_lines_only,
    r#"
        fun main() {
            val text = """
                
                
            """
            println(text.lines().size)
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_raw_string_with_tab_character,
    r#"
        fun main() {
            val text = """a	b"""
            println(text.contains("\t"))
        }
    "#,
    &["false"]
);

kotlin_run_test!(
    test_raw_string_preserves_doublespace,
    r#"
        fun main() {
            val text = """a  b"""
            println(text.replace("  ", "_"))
        }
    "#,
    &["a__b"]
);

kotlin_run_test!(
    test_raw_string_with_dollar_in_margin,
    r#"
        fun main() {
            val text = """
                |$x = 1
            """.trimMargin()
            println(text)
        }
    "#,
    &["$x = 1"]
);

kotlin_run_test!(
    test_raw_string_escaped_newline_not_interpreted,
    r#"
        fun main() {
            val text = """a\nb"""
            println(text.length)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_raw_string_indexed_last,
    r#"
        fun main() {
            val text = """abc"""
            println(text[text.length - 1])
        }
    "#,
    &["c"]
);

kotlin_run_test!(
    test_raw_string_reuse_multiple_times,
    r#"
        fun make(base: String): String = """[$base]"""
        fun main() {
            val x = make("x") + make("y")
            println(x)
        }
    "#,
    &["[x][y]"]
);

kotlin_run_test!(
    test_raw_string_join_lines,
    r#"
        fun main() {
            val text = """a
b"""
            println(text.split('\n').joinToString(","))
        }
    "#,
    &["a,,b"]
);
