kotlin_run_test!(
    test_when_with_guard_first,
    r#"
        fun label(x: Int): String = when {
            x < 0 -> "neg"
            x == 0 -> "zero"
            x % 2 == 0 -> "even"
            else -> "odd"
        }
        fun main() {
            println(label(-1))
            println(label(4))
            println(label(3))
        }
    "#,
    &["neg", "even", "odd"]
);

kotlin_run_test!(
    test_when_with_boolean_guard,
    r#"
        fun isReady(v: Int): Boolean = v > 1 && v < 5
        fun main() {
            val out = when {
                isReady(0) -> "no"
                isReady(3) -> "yes"
                else -> "nope"
            }
            println(out)
        }
    "#,
    &["yes"]
);

kotlin_run_test!(
    test_when_string_length_guard,
    r#"
        fun main() {
            val s = "kotlin"
            val out = when {
                s.isEmpty() -> "empty"
                s.length < 3 -> "tiny"
                s.length == 6 -> "size6"
                else -> "other"
            }
            println(out)
        }
    "#,
    &["size6"]
);

kotlin_run_test!(
    test_when_guard_with_range,
    r#"
        fun score(v: Int): String = when {
            v in 1..3 -> "low"
            v in 4..6 -> "mid"
            v in 7..9 -> "high"
            else -> "out"
        }
        fun main() {
            println(score(2))
            println(score(6))
            println(score(12))
        }
    "#,
    &["low", "mid", "out"]
);

kotlin_run_test!(
    test_when_subject_with_computed_guard,
    r#"
        fun kind(v: Int): String = when (v % 2) {
            0 -> if (v > 0) "even-pos" else "even-neg"
            else -> "odd"
        }
        fun main() {
            println(kind(4))
            println(kind(-3))
        }
    "#,
    &["even-pos", "odd"]
);

kotlin_run_test!(
    test_when_guarded_type,
    r#"
        fun describe(x: Any): String = when {
            x is Int && x > 10 -> "large-int"
            x is Int -> "int"
            x is String && x.isEmpty() -> "empty-str"
            x is String -> "string"
            else -> "other"
        }
        fun main() {
            println(describe(11))
            println(describe(3))
            println(describe(""))
            println(describe("x"))
        }
    "#,
    &["large-int", "int", "empty-str", "string"]
);

kotlin_run_test!(
    test_when_complex_guard,
    r#"
        fun canRun(v: Int): Boolean = when {
            v < 0 -> false
            v == 0 -> false
            v % 2 == 1 -> false
            else -> true
        }
        fun main() {
            println(canRun(2))
            println(canRun(3))
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_when_nested_guard,
    r#"
        fun check(v: Int): String {
            return when {
                v < 0 -> "neg"
                v in 0..5 -> when (v % 2) {
                    0 -> "small-even"
                    else -> "small-odd"
                }
                else -> "large"
            }
        }
        fun main() {
            println(check(2))
            println(check(3))
            println(check(9))
        }
    "#,
    &["small-even", "small-odd", "large"]
);

kotlin_run_test!(
    test_when_reusable_expression,
    r#"
        fun main() {
            val x = 7
            val out = when {
                x > 10 -> "gt"
                x % 2 == 0 -> "even"
                else -> "odd"
            }
            println(out)
        }
    "#,
    &["odd"]
);

kotlin_run_test!(
    test_when_with_local_boolean,
    r#"
        fun main() {
            val enabled = true
            val out = when {
                !enabled -> "off"
                else -> "on"
            }
            println(out)
        }
    "#,
    &["on"]
);

kotlin_run_test!(
    test_when_guarded_array_bounds,
    r#"
        fun main() {
            val value = 3
            val out = when {
                value in 1..3 -> "small"
                value in 4..6 -> "med"
                else -> "large"
            }
            println(out)
        }
    "#,
    &["small"]
);

kotlin_run_test!(
    test_when_guarded_strings,
    r#"
        fun main() {
            val text = ""
            val out = when {
                text.isEmpty() -> "empty"
                text.length < 2 -> "short"
                else -> "ok"
            }
            println(out)
        }
    "#,
    &["empty"]
);

kotlin_run_test!(
    test_when_guarded_null,
    r#"
        fun main() {
            val value: Int? = null
            val out = when {
                value == null -> "null"
                value > 0 -> "pos"
                else -> "other"
            }
            println(out)
        }
    "#,
    &["null"]
);

kotlin_run_test!(
    test_when_guarded_char,
    r#"
        fun main() {
            val c = 'z'
            val out = when {
                c in 'a'..'m' -> "first"
                c in 'n'..'z' -> "last"
                else -> "other"
            }
            println(out)
        }
    "#,
    &["last"]
);

kotlin_run_test!(
    test_when_subject_with_let,
    r#"
        fun classify(x: Int?): String = x?.let {
            when {
                it < 0 -> "neg"
                it == 0 -> "zero"
                else -> "pos"
            }
        } ?: "null"
        fun main() {
            println(classify(-2))
            println(classify(null))
        }
    "#,
    &["neg", "null"]
);

kotlin_run_test!(
    test_when_guarded_boolean_chain,
    r#"
        fun decide(a: Boolean, b: Boolean): String = when {
            a && b -> "both"
            a -> "a"
            b -> "b"
            else -> "none"
        }
        fun main() {
            println(decide(true, true))
            println(decide(true, false))
            println(decide(false, false))
        }
    "#,
    &["both", "a", "none"]
);

kotlin_run_test!(
    test_when_guarded_math,
    r#"
        fun level(v: Int): String = when {
            v * 2 > 10 -> "high"
            v + 1 == 4 -> "four"
            v in 0..2 -> "low"
            else -> "mid"
        }
        fun main() {
            println(level(6))
            println(level(3))
            println(level(1))
        }
    "#,
    &["high", "four", "low"]
);

kotlin_run_test!(
    test_when_guarded_short_circuit,
    r#"
        fun main() {
            val x = 5
            val y = 0
            val out = when {
                x > 0 && y == 0 -> "safe"
                x > 10 && y == 1 -> "skip"
                else -> "other"
            }
            println(out)
        }
    "#,
    &["safe"]
);

kotlin_run_test!(
    test_when_with_nested_guard_scope,
    r#"
        fun main() {
            val threshold = 4
            val value = 3
            val out = when {
                value > threshold -> "too-high"
                else -> {
                    val scaled = value + threshold
                    if (scaled > 5) "scaled" else "small"
                }
            }
            println(out)
        }
    "#,
    &["scaled"]
);

kotlin_run_test!(
    test_when_guarded_multiple_vars,
    r#"
        fun main() {
            val a = 2
            val b = 4
            val out = when {
                a + b > 10 -> "big"
                a * b == 8 -> "match"
                else -> "other"
            }
            println(out)
        }
    "#,
    &["match"]
);

kotlin_run_test!(
    test_when_guard_fallback,
    r#"
        fun main() {
            val v = -1
            val out = when {
                v > 10 -> "high"
                v < 0 -> "low"
                else -> "mid"
            }
            println(out)
        }
    "#,
    &["low"]
);

kotlin_run_test!(
    test_when_guarded_char_type,
    r#"
        fun toCategory(c: Char): String = when {
            c == 'x' || c == 'y' -> "xy"
            c in 'a'..'f' -> "alpha"
            c.isDigit() -> "digit"
            else -> "other"
        }
        fun main() {
            println(toCategory('x'))
            println(toCategory('b'))
            println(toCategory('7'))
        }
    "#,
    &["xy", "alpha", "digit"]
);

kotlin_run_test!(
    test_when_guarded_by_length,
    r#"
        fun label(s: String): String = when {
            s.length == 0 -> "empty"
            s.length == 1 -> "tiny"
            s.length > 3 -> "long"
            else -> "short"
        }
        fun main() {
            println(label(""))
            println(label("a"))
            println(label("code"))
        }
    "#,
    &["empty", "tiny", "long"]
);

kotlin_run_test!(
    test_when_guarded_fallback_order,
    r#"
        fun label(v: Int): String = when {
            v < 0 -> "neg"
            v % 2 == 0 -> "even"
            else -> "odd"
        }
        fun main() {
            println(label(-2))
            println(label(3))
        }
    "#,
    &["neg", "odd"]
);

kotlin_run_test!(
    test_when_guarded_with_function_call,
    r#"
        fun isEven(v: Int): Boolean = v % 2 == 0
        fun label(v: Int): String = when {
            isEven(v) && v > 0 -> "positive-even"
            v > 0 -> "positive"
            else -> "not"
        }
        fun main() {
            println(label(4))
            println(label(3))
        }
    "#,
    &["positive-even", "positive"]
);

kotlin_run_test!(
    test_when_guarded_chain,
    r#"
        fun label(a: Int, b: Int): String = when {
            a == 0 || b == 0 -> "zero"
            a == b -> "equal"
            a + b == 10 -> "ten"
            else -> "other"
        }
        fun main() {
            println(label(0, 5))
            println(label(5, 5))
            println(label(3, 7))
        }
    "#,
    &["zero", "equal", "ten"]
);

kotlin_run_test!(
    test_when_guarded_nullable,
    r#"
        fun label(v: Int?): String = when {
            v == null -> "null"
            v > 3 -> "big"
            else -> "small"
        }
        fun main() {
            println(label(null))
            println(label(2))
        }
    "#,
    &["null", "small"]
);

kotlin_run_test!(
    test_when_guarded_with_chars,
    r#"
        fun label(c: Char): String = when {
            c in 'a'..'f' -> "alpha"
            c in 'g'..'m' -> "middle"
            c in 'n'..'z' -> "late"
            else -> "other"
        }
        fun main() {
            println(label('c'))
            println(label('k'))
            println(label('z'))
        }
    "#,
    &["alpha", "middle", "late"]
);

kotlin_run_test!(
    test_when_guarded_with_types,
    r#"
        fun label(v: Any): String = when {
            v is String && v.isEmpty() -> "empty"
            v is String -> "str"
            v is Int && v > 10 -> "big-int"
            v is Int -> "int"
            else -> "other"
        }
        fun main() {
            println(label(""))
            println(label("x"))
            println(label(11))
            println(label(5))
        }
    "#,
    &["empty", "str", "big-int", "int"]
);

kotlin_run_test!(
    test_when_guarded_in_nested_block,
    r#"
        fun main() {
            val value = 7
            val out = when {
                value < 0 -> "neg"
                value < 5 -> "low"
                else -> {
                    if (value % 2 == 1) "odd" else "even"
                }
            }
            println(out)
        }
    "#,
    &["odd"]
);

kotlin_run_test!(
    test_when_guarded_boolean_math,
    r#"
        fun main() {
            val out = when {
                1 + 1 == 3 -> "wrong"
                1 + 1 == 2 -> "yes"
                else -> "no"
            }
            println(out)
        }
    "#,
    &["yes"]
);
