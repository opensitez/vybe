kotlin_run_test!(
    test_when_int_literal_branches,
    r#"
        fun valueName(x: Int): String = when (x) {
            0 -> "zero"
            1 -> "one"
            2 -> "two"
            else -> "other"
        }
        fun main() {
            println(valueName(0))
            println(valueName(3))
        }
    "#,
    &["zero", "other"]
);

kotlin_run_test!(
    test_when_range_membership,
    r#"
        fun inRange(x: Int): String = when (x) {
            in 1..3 -> "small"
            in 4..6 -> "mid"
            else -> "big"
        }
        fun main() {
            println(inRange(2))
            println(inRange(5))
            println(inRange(9))
        }
    "#,
    &["small", "mid", "big"]
);

kotlin_run_test!(
    test_when_in_keyword,
    r#"
        fun classify(ch: Char): String = when (ch) {
            'a', 'b', 'c' -> "abc"
            'x', 'y', 'z' -> "xyz"
            else -> "other"
        }
        fun main() {
            println(classify('b'))
            println(classify('z'))
        }
    "#,
    &["abc", "xyz"]
);

kotlin_run_test!(
    test_when_type_check,
    r#"
        fun tag(v: Any): String = when (v) {
            is Int -> "int"
            is String -> "string"
            else -> "other"
        }
        fun main() {
            println(tag(5))
            println(tag("x"))
            println(tag(1.5))
        }
    "#,
    &["int", "string", "other"]
);

kotlin_run_test!(
    test_when_without_subject,
    r#"
        fun main() {
            val v = 7
            val out = when {
                v < 0 -> "neg"
                v < 5 -> "low"
                v == 7 -> "seven"
                else -> "other"
            }
            println(out)
        }
    "#,
    &["seven"]
);

kotlin_run_test!(
    test_when_guard_condition,
    r#"
        fun main() {
            val p = Pair(2, 4)
            val out = when {
                p.first == p.second -> "eq"
                p.first + p.second == 6 -> "sum-six"
                else -> "other"
            }
            println(out)
        }
    "#,
    &["sum-six"]
);

kotlin_run_test!(
    test_when_subject_with_function,
    r#"
        fun isGood(x: Int): Boolean = x % 2 == 0
        fun status(x: Int): String = when (x) {
            0 -> "zero"
            else -> if (isGood(x)) "even" else "odd"
        }
        fun main() {
            println(status(4))
            println(status(5))
        }
    "#,
    &["even", "odd"]
);

kotlin_run_test!(
    test_when_returns_expression,
    r#"
        fun map(v: Int): Int = when (v) {
            1 -> 10
            2 -> 20
            else -> v
        }
        fun main() {
            println(map(1))
            println(map(3))
        }
    "#,
    &["10", "3"]
);

kotlin_run_test!(
        test_when_in_set_or_range,
    r#"
        fun classify(x: Int): String = when (x) {
            1, 2, 3 -> "small"
            in 4..6 -> "mid"
            else -> "other"
        }
        fun main() {
            println(classify(2))
            println(classify(5))
            println(classify(10))
        }
    "#,
    &["small", "mid", "other"]
);

kotlin_run_test!(
    test_when_subject_reassignment,
    r#"
        fun main() {
            var x = 1
            val out = when (x) {
                1 -> {
                    x = 2
                    "one"
                }
                else -> "other"
            }
            println(out)
            println(x)
        }
    "#,
    &["one", "2"]
);

kotlin_run_test!(
    test_when_boolean,
    r#"
        fun resolve(x: Boolean): String = when (x) {
            true -> "yes"
            false -> "no"
        }
        fun main() {
            println(resolve(true))
            println(resolve(false))
        }
    "#,
    &["yes", "no"]
);

kotlin_run_test!(
    test_when_subject_object_type,
    r#"
        open class Animal
        class Dog : Animal()
        class Cat : Animal()
        fun identify(a: Animal): String = when (a) {
            is Dog -> "dog"
            is Cat -> "cat"
            else -> "other"
        }
        fun main() {
            println(identify(Dog()))
            println(identify(Cat()))
        }
    "#,
    &["dog", "cat"]
);

kotlin_run_test!(
    test_when_empty_else,
    r#"
        fun label(x: Int): String = when (x) {
            0 -> "zero"
            else -> "not"
        }
        fun main() {
            println(label(0))
            println(label(1))
        }
    "#,
    &["zero", "not"]
);

kotlin_run_test!(
    test_when_with_multiple_checks,
    r#"
        fun classify(v: Int): String = when (v) {
            1, 2, 3 -> "low"
            4, 5, 6 -> "mid"
            else -> "high"
        }
        fun main() {
            println(classify(2))
            println(classify(5))
            println(classify(9))
        }
    "#,
    &["low", "mid", "high"]
);

kotlin_run_test!(
    test_when_nested_when,
    r#"
        fun main() {
            val a = 5
            val out = when (a) {
                in 1..10 -> when (a % 2) {
                    0 -> "even"
                    else -> "odd"
                }
                else -> "none"
            }
            println(out)
        }
    "#,
    &["odd"]
);

kotlin_run_test!(
    test_when_string_subject,
    r#"
        fun describe(input: String): String = when (input.length) {
            0 -> "empty"
            in 1..3 -> "short"
            else -> "long"
        }
        fun main() {
            println(describe(""))
            println(describe("ok"))
            println(describe("hello"))
        }
    "#,
    &["empty", "short", "long"]
);

kotlin_run_test!(
    test_when_subject_local_let,
    r#"
        fun main() {
            val x = 10
            val out = when (x) {
                5 -> 1
                10 -> {
                    val y = x / 2
                    y
                }
                else -> 0
            }
            println(out)
        }
    "#,
    &["5"]
);

kotlin_run_test!(
    test_when_subject_block_with_side_effect,
    r#"
        var seen = 0
        fun classify(x: Int): String {
            return when (x) {
                1 -> { seen = 1; "one" }
                2 -> { seen = 2; "two" }
                else -> { seen = 3; "other" }
            }
        }
        fun main() {
            println(classify(2))
            println(seen)
        }
    "#,
    &["two", "2"]
);

kotlin_run_test!(
    test_when_subject_default_for_unknown,
    r#"
        fun status(x: String): String = when (x) {
            "on" -> "1"
            "off" -> "0"
            else -> "x"
        }
        fun main() {
            println(status("on"))
            println(status("pause"))
        }
    "#,
    &["1", "x"]
);

kotlin_run_test!(
    test_when_subject_with_is_checks,
    r#"
        fun describe(x: Any): String = when (x) {
            is Int -> "int"
            is Double -> "double"
            is String -> "string"
            else -> "other"
        }
        fun main() {
            println(describe("x"))
            println(describe(4.5))
        }
    "#,
    &["string", "double"]
);

kotlin_run_test!(
        test_when_subject_and_guarded_in_array,
    r#"
        fun classify(x: Int): String = when (x) {
            2, 4, 6 -> "even-basic"
            in 7..9 -> "high-seven-nine"
            else -> "other"
        }
        fun main() {
            println(classify(2))
            println(classify(8))
            println(classify(12))
        }
    "#,
    &["even-basic", "high-seven-nine", "other"]
);

kotlin_run_test!(
    test_when_multiple_subject_types,
    r#"
        fun classify(v: Any): String = when (v) {
            is Int -> "number"
            is Char -> "char"
            is Boolean -> "bool"
            else -> "other"
        }
        fun main() {
            println(classify('a'))
            println(classify(false))
        }
    "#,
    &["char", "bool"]
);

kotlin_run_test!(
        test_when_inline_subject,
    r#"
        fun main() {
            val n = 8
            val result = when (n) {
                1 -> "one"
                8 -> "eight"
                else -> "other"
            }
            println(result)
        }
    "#,
    &["eight"]
);

kotlin_run_test!(
    test_when_fallback_with_expression,
    r#"
        fun score(x: Int): Int = when (x) {
            in 1..3 -> 1
            in 4..6 -> 2
            else -> 3
        }
        fun main() {
            println(score(4))
            println(score(9))
        }
    "#,
    &["2", "3"]
);

kotlin_run_test!(
        test_when_subject_order,
    r#"
        fun pick(x: Int): String = when (x) {
            in 1..5 -> "low"
            6 -> "six"
            in 6..9 -> "mid"
            else -> "high"
        }
        fun main() {
            println(pick(2))
            println(pick(6))
            println(pick(8))
        }
    "#,
    &["low", "six", "mid"]
);

kotlin_run_test!(
    test_when_subject_with_null,
    r#"
        fun safeDescribe(v: Int?): String = when (v) {
            null -> "null"
            0 -> "zero"
            else -> "other"
        }
        fun main() {
            println(safeDescribe(null))
            println(safeDescribe(0))
        }
    "#,
    &["null", "zero"]
);

kotlin_run_test!(
    test_when_subject_with_negatives,
    r#"
        fun main() {
            val x = -3
            val out = when (x) {
                in Int.MIN_VALUE..-1 -> "neg"
                in 0..9 -> "small"
                else -> "other"
            }
            println(out)
        }
    "#,
    &["neg"]
);

kotlin_run_test!(
    test_when_subject_with_boolean,
    r#"
        fun main() {
            val x = true
            val out = when (x) {
                true -> "ok"
                false -> "no"
            }
            println(out)
        }
    "#,
    &["ok"]
);

kotlin_run_test!(
    test_when_subject_with_fallback_pair,
    r#"
        fun main() {
            val pair = Pair(9, 1)
            val out = when (pair) {
                Pair(1, 1) -> "11"
                Pair(9, 1) -> "91"
                else -> "other"
            }
            println(out)
        }
    "#,
    &["91"]
);

kotlin_run_test!(
    test_when_subject_reassigning,
    r#"
        fun main() {
            var value = 4
            val out = when (value) {
                4 -> {
                    value += 1
                    "four"
                }
                else -> "other"
            }
            println(out)
            println(value)
        }
    "#,
    &["four", "5"]
);
