kotlin_run_test!(
    test_string_length_and_indices,
    r#"
        fun main() {
            val text = "kotlin"
            println(text.length)
            println(text[text.length - 1])
        }
    "#,
    &["6", "n"]
);

kotlin_run_test!(
    test_string_contains_and_starts_with,
    r#"
        fun main() {
            val text = "language"
            println(text.startsWith("lang"))
            println(text.contains("gua"))
            println(text.endsWith("age"))
        }
    "#,
    &["true", "true", "true"]
);

kotlin_run_test!(
    test_string_substring_slices,
    r#"
        fun main() {
            val text = "kotlin-lang"
            println(text.substring(0, 6))
            println(text.substring(7))
        }
    "#,
    &["kotlin", "lang"]
);

kotlin_run_test!(
    test_string_replace_and_index_of,
    r#"
        fun main() {
            val text = "aa-bb-aa"
            println(text.replace("aa", "x"))
            println(text.indexOf("bb"))
        }
    "#,
    &["x-bb-x", "3"]
);

kotlin_run_test!(
    test_string_split_and_join,
    r#"
        fun main() {
            val text = "a,b,c"
            println(text.split(",").joinToString("|"))
            println(text.reversed())
        }
    "#,
    &["a|b|c", "c,b,a"]
);

kotlin_run_test!(
    test_string_trim_and_compare,
    r#"
        fun main() {
            val text = "  Kotlin  "
            println(text.trim())
            println(text.trim().lowercase() == "kotlin")
        }
    "#,
    &["Kotlin", "true"]
);

kotlin_run_test!(
    test_string_take_drop,
    r#"
        fun main() {
            val text = "abcdef"
            println(text.take(2))
            println(text.drop(3))
            println(text.dropLast(2))
        }
    "#,
    &["ab", "def", "abcd"]
);

kotlin_run_test!(
    test_string_repeat_and_padding,
    r#"
        fun main() {
            val text = "x"
            println(text.repeat(3))
            println("1".padStart(3, '0'))
            println("7".padEnd(3, '0'))
        }
    "#,
    &["xxx", "001", "700"]
);

kotlin_run_test!(
    test_string_template_escaping,
    r#"
        fun main() {
            val name = "k"
            val count = 2
            println("$name$count")
            println("${name.uppercase()}$count")
        }
    "#,
    &["k2", "K2"]
);

kotlin_run_test!(
    test_string_splitter_with_limit,
    r#"
        fun main() {
            val text = "a|b|c|d"
            val pieces = text.split("|", limit = 2)
            println(pieces.size)
            println(pieces[0])
            println(pieces[1])
        }
    "#,
    &["2", "a", "b|c|d"]
);

kotlin_run_test!(
    test_string_is_blank_and_empty,
    r#"
        fun main() {
            println("".isEmpty())
            println("   ".isBlank())
            println("x".isBlank())
        }
    "#,
    &["true", "true", "false"]
);
