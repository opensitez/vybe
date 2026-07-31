kotlin_run_test!(
    test_in_range_membership,
    r#"
        fun main() {
            println(3 in 1..5)
            println(7 in 1..5)
            println(5 !in 1 until 5)
        }
    "#,
    &["true", "false", "true"]
);

kotlin_run_test!(
    test_in_list_membership,
    r#"
        fun main() {
            val values = listOf("a", "b", "c")
            println("b" in values)
            println("z" !in values)
        }
    "#,
    &["true", "true"]
);

kotlin_run_test!(
    test_in_set_lookup,
    r#"
        fun main() {
            val values = setOf(1, 2, 3)
            println(2 in values)
            println(8 in values)
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_in_string_char_membership,
    r#"
        fun main() {
            val text = "kotlin"
            println('o' in text)
            println('x' !in text)
        }
    "#,
    &["true", "true"]
);

kotlin_run_test!(
    test_in_map_contains_key,
    r#"
        fun main() {
            val map = mapOf("a" to 1, "b" to 2)
            println("a" in map)
            println("c" in map)
            println("a" !in map)
        }
    "#,
    &["true", "false", "false"]
);

kotlin_run_test!(
    test_in_range_exclusive_until,
    r#"
        fun main() {
            println(5 in 1 until 5)
            println(4 in 1 until 5)
            println(0 until 5)
            println(3 in 1..4)
        }
    "#,
    &["false", "true", "1,2,3,4", "true"]
);

kotlin_run_test!(
    test_in_with_down_to_range,
    r#"
        fun main() {
            println(3 in 5 downTo 1)
            println(6 in 5 downTo 1)
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_in_range_with_step,
    r#"
        fun main() {
            println(5 in 1..10 step 2)
            println(6 in 1..10 step 2)
            println(4 in 10 downTo 1 step 2)
            println(5 in 10 downTo 1 step 2)
        }
    "#,
    &["true", "false", "true", "false"]
);

kotlin_run_test!(
    test_in_with_if_condition,
    r#"
        fun classify(v: Int): String {
            return if (v in 1..3) "small" else if (v in 4..6) "mid" else "big"
        }

        fun main() {
            println(classify(2))
            println(classify(5))
            println(classify(7))
        }
    "#,
    &["small", "mid", "big"]
);

kotlin_run_test!(
    test_in_subjectless_when_with_is_and_in,
    r#"
        fun label(v: Int): String {
            return when {
                v in 1..2 -> "low"
                v in 3..4 -> "mid"
                else -> "high"
            }
        }

        fun main() {
            println(label(2))
            println(label(4))
            println(label(7))
        }
    "#,
    &["low", "mid", "high"]
);
