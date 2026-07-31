kotlin_run_test!(
    test_multiplication_precedes_addition,
    r#"
        fun main() {
            println(2 + 3 * 4)
        }
    "#,
    &["14"]
);

kotlin_run_test!(
    test_parenthesized_changes_addition_precedence,
    r#"
        fun main() {
            println((2 + 3) * 4)
        }
    "#,
    &["20"]
);

kotlin_run_test!(
    test_subtraction_left_assoc,
    r#"
        fun main() {
            println(20 - 5 - 3)
        }
    "#,
    &["12"]
);

kotlin_run_test!(
    test_mixed_arithmetic_precedence,
    r#"
        fun main() {
            println(10 + 2 * 3 - 4 / 2)
        }
    "#,
    &["14"]
);

kotlin_run_test!(
    test_boolean_and_before_or,
    r#"
        fun main() {
            println(true || false && false)
        }
    "#,
    &["true"]
);

kotlin_run_test!(
    test_boolean_or_before_compare,
    r#"
        fun main() {
            println((false || true) && false)
        }
    "#,
    &["false"]
);

kotlin_run_test!(
    test_comparison_precedence,
    r#"
        fun main() {
            println(3 + 2 > 4 && 2 * 2 == 4)
        }
    "#,
    &["true"]
);

kotlin_run_test!(
    test_elvis_with_addition,
    r#"
        fun main() {
            val value: Int? = null
            println((value ?: 5) + 7)
        }
    "#,
    &["12"]
);

kotlin_run_test!(
    test_not_before_equality,
    r#"
        fun main() {
            val value = 3
            println(! (value == 4))
        }
    "#,
    &["true"]
);

kotlin_run_test!(
    test_string_interpolation_order,
    r#"
        fun main() {
            val a = 1
            val b = 2
            val c = 3
            println(a + b * c)
            println("${a + b}*${c}")
        }
    "#,
    &["7", "3*3"]
);
