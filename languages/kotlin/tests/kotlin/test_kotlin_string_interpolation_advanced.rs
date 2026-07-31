kotlin_run_test!(
    test_nested_expression_interpolation,
    r#"
        fun main() {
            val value = 5
            val msg = "value is ${'$'}{if (value > 3) "high" else "low"} and ${'$'}value"
            println(msg)
        }
    "#,
    &["value is high and 5"]
);

kotlin_run_test!(
    test_string_template_with_function_call,
    r#"
        fun decorate(input: String): String = input.uppercase()
        fun main() {
            println("${'$'}{decorate("kotlin")}")
        }
    "#,
    &["KOTLIN"]
);

kotlin_run_test!(
    test_interpolation_in_loop_accumulation,
    r#"
        fun main() {
            var total = 0
            for (i in 1..4) {
                total += i
            }
            println("sum=${'$'}total")
        }
    "#,
    &["sum=10"]
);

kotlin_run_test!(
    test_interpolation_with_raw_indexed_access,
    r#"
        fun main() {
            val values = listOf("ab", "cd", "ef")
            println("first=${'$'}{values[0]} len=${'$'}{values[0].length}")
        }
    "#,
    &["first=ab len=2"]
);

kotlin_run_test!(
    test_escaped_dollar_output,
    r#"
        fun main() {
            println("literal ${'$'}")
            println("price ${'$'}{10}")
        }
    "#,
    &["literal $", "price 10"]
);

kotlin_run_test!(
    test_interpolation_with_boolean_logic,
    r#"
        fun main() {
            val ok = true
            println("state=${'$'}{ok && true}")
        }
    "#,
    &["state=true"]
);

kotlin_run_test!(
    test_interpolation_with_nullable_value,
    r#"
        fun label(value: String?): String {
            return "${'$'}{value ?: "none"}"
        }
        fun main() {
            println(label(null))
            println(label("x"))
        }
    "#,
    &["none", "x"]
);

kotlin_run_test!(
    test_multi_line_template_concatenation,
    r#"
        fun main() {
            val lines = """
                a
                b
            """.trimIndent()
            println("${'$'}{lines.lines().size}")
            println(lines[0])
        }
    "#,
    &["2", "a"]
);

kotlin_run_test!(
    test_interpolation_without_braces,
    r#"
        fun main() {
            val count = 4
            println("${'$'}count items")
        }
    "#,
    &["4 items"]
);

kotlin_run_test!(
    test_interpolation_with_local_function_call,
    r#"
        fun value(a: Int) = a * 2
        fun main() {
            val x = 3
            println("doubled=${'$'}{value(x)}")
        }
    "#,
    &["doubled=6"]
);

kotlin_run_test!(
    test_interpolation_of_char_code_and_codepoint,
    r#"
        fun main() {
            val ch = 'A'
            println("${'$'}ch-${'$'}{ch.code}")
        }
    "#,
    &["A-65"]
);
