kotlin_run_test!(
    test_eq_true,
    r#"
        fun main() {
            println(1 == 1)
            println(1 == 2)
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_neq,
    r#"
        fun main() {
            println(1 != 2)
            println(1 != 1)
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_less_than,
    r#"
        fun main() {
            println(1 < 2)
            println(2 < 2)
            println(3 < 2)
        }
    "#,
    &["true", "false", "false"]
);

kotlin_run_test!(
    test_greater_than,
    r#"
        fun main() {
            println(3 > 2)
            println(2 > 2)
            println(1 > 4)
        }
    "#,
    &["true", "false", "false"]
);

kotlin_run_test!(
    test_leq,
    r#"
        fun main() {
            println(2 <= 2)
            println(1 <= 2)
            println(3 <= 2)
        }
    "#,
    &["true", "true", "false"]
);

kotlin_run_test!(
    test_geq,
    r#"
        fun main() {
            println(2 >= 2)
            println(3 >= 2)
            println(1 >= 2)
        }
    "#,
    &["true", "true", "false"]
);

kotlin_run_test!(
    test_compare_with_booleans,
    r#"
        fun main() {
            val a = true
            val b = false
            println(a == b)
            println(a != b)
        }
    "#,
    &["false", "true"]
);

kotlin_run_test!(
    test_string_equality,
    r#"
        fun main() {
            println("a" == "a")
            println("a" == "b")
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_reference_equality_not_eq,
    r#"
        fun main() {
            val a: Any = "x"
            val b: Any = "x"
            println(a == b)
            println(a === b)
        }
    "#,
    &["true", "true"]
);

kotlin_run_test!(
    test_reference_inequality,
    r#"
        fun main() {
            val a = Any()
            val b = Any()
            println(a === b)
        }
    "#,
    &["false"]
);

kotlin_run_test!(
    test_compare_int_and_long,
    r#"
        fun main() {
            println(1L > 0)
            println(1 == 1L)
        }
    "#,
    &["true", "true"]
);

kotlin_run_test!(
    test_compare_char,
    r#"
        fun main() {
            println('a' < 'c')
            println('z' > 'm')
            println('a' == 'a')
        }
    "#,
    &["true", "true", "true"]
);

kotlin_run_test!(
    test_compare_float_nans,
    r#"
        fun main() {
            val x = 0.0 / 0.0
            println(x == x)
            println(x != x)
        }
    "#,
    &["false", "true"]
);

kotlin_run_test!(
    test_compare_with_ranges,
    r#"
        fun main() {
            val r = 1..4
            println(2 in r)
            println(5 in r)
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_compare_with_until,
    r#"
        fun main() {
            val r = 1 until 4
            println(3 in r)
            println(4 in r)
            println(0 in r)
        }
    "#,
    &["true", "false", "false"]
);

kotlin_run_test!(
    test_compare_chained,
    r#"
        fun main() {
            val x = 2
            println(x > 1 && x < 5)
            println(x < 2 || x > 10)
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_compare_nested,
    r#"
        fun main() {
            println((1 < 2) == true)
            println((1 < 2) == false)
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_compare_tuples,
    r#"
        data class P(val x: Int, val y: Int)
        fun main() {
            println(P(1,2) == P(1,2))
            println(P(1,2) == P(2,1))
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_compare_strings_length,
    r#"
        fun main() {
            val a = "abc"
            val b = "def"
            println(a.length == b.length)
            println(a.length < b.length)
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_compare_array_size,
    r#"
        fun main() {
            val a = intArrayOf(1,2,3)
            val b = intArrayOf(1,2)
            println(a.size > b.size)
            println(a.size == b.size)
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_compare_ternary_chain,
    r#"
        fun main() {
            val out = if (1 < 2) if (2 < 3) "yes" else "no" else "no"
            println(out)
        }
    "#,
    &["yes"]
);

kotlin_run_test!(
    test_compare_sign_inversion,
    r#"
        fun isPositive(n: Int): Boolean {
            return n > 0
        }
        fun main() {
            println(isPositive(-1) == !isPositive(1))
            println(!isPositive(0) == isPositive(-1))
        }
    "#,
    &["true", "true"]
);

kotlin_run_test!(
    test_compare_nullable,
    r#"
        fun main() {
            val a: Int? = null
            println(a == null)
            println(a != null)
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_compare_zero_division,
    r#"
        fun main() {
            try {
                println(1 == (1 / 0))
            } catch (e: Exception) {
                println("err")
            }
        }
    "#,
    &["err"]
);

kotlin_run_test!(
    test_compare_ordered_pairs,
    r#"
        fun cmp(a: Int, b: Int): String {
            return if (a == b) "equal" else if (a < b) "lt" else "gt"
        }
        fun main() {
            println(cmp(1, 1))
            println(cmp(2, 4))
            println(cmp(7, 3))
        }
    "#,
    &["equal", "lt", "gt"]
);

kotlin_run_test!(
    test_compare_when,
    r#"
        fun main() {
            val x = 4
            val result = when {
                x < 2 -> "small"
                x == 4 -> "four"
                else -> "other"
            }
            println(result)
        }
    "#,
    &["four"]
);

kotlin_run_test!(
    test_compare_object_identity,
    r#"
        class Item
        fun main() {
            val a = Item()
            val b = Item()
            val c = a
            println(a === b)
            println(a === c)
        }
    "#,
    &["false", "true"]
);

kotlin_run_test!(
    test_compare_long_overflow_bound,
    r#"
        fun main() {
            val a = Long.MAX_VALUE
            val b = a + 1
            println(a < b)
            println(a == b)
        }
    "#,
    &["false", "false"]
);

kotlin_run_test!(
    test_compare_nested_ranges,
    r#"
        fun main() {
            val x = 5
            val y = x in 1..10
            val z = x !in 10..20
            println(y)
            println(z)
        }
    "#,
    &["true", "true"]
);

kotlin_run_test!(
    test_compare_with_char_codes,
    r#"
        fun main() {
            println('a'.code < 'b'.code)
            println('z'.code > 'y'.code)
        }
    "#,
    &["true", "true"]
);
