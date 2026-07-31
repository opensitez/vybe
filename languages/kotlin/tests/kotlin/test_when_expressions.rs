use crate::helpers::run_prints;

#[test]
fn test_when_with_expression_subject_returns_matching_branch() {
    let out = run_prints(r#"
        fun score(label: Int): String {
            return when (label) {
                0 -> "zero"
                1 -> "one"
                2, 3 -> "small"
                else -> "many"
            }
        }

        fun main() {
            println(score(0))
            println(score(2))
            println(score(7))
        }
    "#);
    assert_eq!(out, &["zero", "small", "many"]);
}

#[test]
fn test_when_with_multiple_subject_expressions() {
    let out = run_prints(r#"
        fun classify(value: Int): String {
            return when (value) {
                in 1..3 -> "low"
                in 4..10 -> "mid"
                !in 1..10 -> "outside"
                else -> "other"
            }
        }

        fun main() {
            println(classify(2))
            println(classify(10))
            println(classify(20))
        }
    "#);
    assert_eq!(out, &["low", "mid", "outside"]);
}

#[test]
fn test_when_with_guard_condition() {
    let out = run_prints(r#"
        fun tag(value: Int): String {
            return when {
                value > 10 -> "gt"
                value == 10 -> "eq"
                else -> "lt"
            }
        }

        fun main() {
            println(tag(12))
            println(tag(10))
            println(tag(2))
        }
    "#);
    assert_eq!(out, &["gt", "eq", "lt"]);
}

#[test]
fn test_when_with_type_checks() {
    let out = run_prints(r#"
        fun classify(value: Any): String {
            return when (value) {
                is Int -> "int"
                is String -> "string"
                else -> "other"
            }
        }

        fun main() {
            println(classify(1))
            println(classify("x"))
            println(classify(2.0))
        }
    "#);
    assert_eq!(out, &["int", "string", "other"]);
}

#[test]
fn test_when_subject_evaluates_once_with_side_effects() {
    let out = run_prints(r#"
        var ticks = 0

        fun next(): Int {
            ticks += 1
            return ticks
        }

        fun classify(): Int {
            return when (next()) {
                1 -> 10
                2 -> 20
                3 -> 30
                else -> 40
            }
        }

        fun main() {
            println(classify())
            println(classify())
            println(ticks)
        }
    "#);
    assert_eq!(out, &["10", "20", "2"]);
}

#[test]
fn test_when_nested_scoping_and_binding() {
    let out = run_prints(r#"
        fun describe(a: Int, b: Int): String {
            return when (a) {
                0 -> when (b) {
                    0 -> "a0b0"
                    else -> "a0bN"
                }
                else -> when {
                    b == 0 -> "aNb0"
                    b > 10 -> "aNbH"
                    else -> "aNbL"
                }
            }
        }

        fun main() {
            println(describe(0, 0))
            println(describe(0, 4))
            println(describe(5, 12))
        }
    "#);
    assert_eq!(out, &["a0b0", "a0bN", "aNbH"]);
}

#[test]
fn test_when_with_variable_subject_binding() {
    let out = run_prints(r#"
        fun main() {
            val value = 7
            val label = when (value) {
                is Int -> "int-" + value
                else -> "none"
            }
            println(label)
        }
    "#);
    assert_eq!(out, &["int-7"]);
}

#[test]
fn test_when_reduces_on_collection_size() {
    let out = run_prints(r#"
        fun sizeLabel(values: List<Int>): String {
            return when (values.size) {
                0 -> "empty"
                in 1..2 -> "small"
                in 3..4 -> "mid"
                else -> "large"
            }
        }

        fun main() {
            println(sizeLabel(listOf()))
            println(sizeLabel(listOf(1)))
            println(sizeLabel(listOf(1, 2, 3)))
            println(sizeLabel(listOf(1, 2, 3, 4, 5)))
        }
    "#);
    assert_eq!(out, &["empty", "small", "mid", "large"]);
}

#[test]
fn test_when_with_non_exhaustive_else_on_any_subject() {
    let out = run_prints(r#"
        fun main() {
            println(when ("x") {
                "a" -> 1
                "b" -> 2
                else -> 3
            })
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_when_as_standalone_statement_for_side_effects() {
    let out = run_prints(r#"
        fun main() {
            var acc = ""
            val value = 4
            when {
                value > 10 -> acc = "big"
                value > 1 -> acc = "mid"
                else -> acc = "small"
            }
            println(acc)
        }
    "#);
    assert_eq!(out, &["mid"]);
}

#[test]
fn test_when_with_multiple_conditions_same_branch() {
    let out = run_prints(r#"
        fun classify(value: Int): String {
            return when (value) {
                1, 2, 3 -> "low"
                4, 5, 6 -> "mid"
                else -> "high"
            }
        }

        fun main() {
            println(classify(2))
            println(classify(6))
            println(classify(9))
        }
    "#);
    assert_eq!(out, &["low", "mid", "high"]);
}

#[test]
fn test_when_with_subject_as_nullable_and_null_branch() {
    let out = run_prints(r#"
        fun label(value: String?): String {
            return when (value) {
                null -> "null"
                "" -> "empty"
                else -> "value"
            }
        }

        fun main() {
            println(label(null))
            println(label(""))
            println(label("ok"))
        }
    "#);
    assert_eq!(out, &["null", "empty", "value"]);
}

#[test]
fn test_when_with_guard_first_match_is_used() {
    let out = run_prints(r#"
        fun check(value: Int): String {
            return when {
                value > 0 && value % 2 == 0 -> "positive-even"
                value > 0 -> "positive-odd"
                value < 0 -> "negative"
                else -> "zero"
            }
        }

        fun main() {
            println(check(6))
            println(check(5))
            println(check(-3))
            println(check(0))
        }
    "#);
    assert_eq!(out, &["positive-even", "positive-odd", "negative", "zero"]);
}

#[test]
fn test_when_with_boolean_subject() {
    let out = run_prints(r#"
        fun label(flag: Boolean): String {
            return when (flag) {
                true -> "on"
                false -> "off"
            }
        }

        fun main() {
            println(label(true))
            println(label(false))
        }
    "#);
    assert_eq!(out, &["on", "off"]);
}

#[test]
fn test_when_with_string_subject_patterns() {
    let out = run_prints(r#"
        fun bucket(word: String): String {
            return when (word.lowercase()) {
                "yes", "y", "oui" -> "affirmative"
                "no", "n", "non" -> "negative"
                else -> "unknown"
            }
        }

        fun main() {
            println(bucket("YES"))
            println(bucket("No"))
            println(bucket("maybe"))
        }
    "#);
    assert_eq!(out, &["affirmative", "negative", "unknown"]);
}

#[test]
fn test_when_with_char_subject() {
    let out = run_prints(r#"
        fun kind(ch: Char): String {
            return when (ch) {
                in 'a'..'f' -> "low"
                in 'g'..'m' -> "mid"
                in 'n'..'z' -> "high"
                else -> "other"
            }
        }

        fun main() {
            println(kind('c'))
            println(kind('h'))
            println(kind('x'))
            println(kind('2'))
        }
    "#);
    assert_eq!(out, &["low", "mid", "high", "other"]);
}

#[test]
fn test_when_subject_capture_with_is_and_casts() {
    let out = run_prints(r#"
        interface Shape { fun kind(): String }
        class Dot : Shape {
            override fun kind(): String = "dot"
        }
        class Box(val size: Int) : Shape {
            override fun kind(): String = "box:" + size
        }

        fun describe(shape: Shape): String {
            return when (shape) {
                is Dot -> "dot"
                is Box -> "box=" + shape.size
                else -> "unknown"
            }
        }

        fun main() {
            println(describe(Dot()))
            println(describe(Box(7)))
        }
    "#);
    assert_eq!(out, &["dot", "box=7"]);
}

#[test]
fn test_when_with_inclusive_ranges_and_open_edges() {
    let out = run_prints(r#"
        fun level(value: Int): String {
            return when (value) {
                Int.MIN_VALUE..-1 -> "negative"
                0 -> "zero"
                1..99 -> "low"
                100..Int.MAX_VALUE -> "high"
                else -> "other"
            }
        }

        fun main() {
            println(level(-4))
            println(level(0))
            println(level(1))
            println(level(100))
        }
    "#);
    assert_eq!(out, &["negative", "zero", "low", "high"]);
}

#[test]
fn test_when_with_list_subject_and_contains_style_check() {
    let out = run_prints(r#"
        fun classify(value: Int): String {
            return when (value) {
                in listOf(1, 3, 5) -> "odd-primeish"
                in listOf(2, 4, 6) -> "even-small"
                else -> "other"
            }
        }

        fun main() {
            println(classify(1))
            println(classify(4))
            println(classify(7))
        }
    "#);
    assert_eq!(out, &["odd-primeish", "even-small", "other"]);
}

#[test]
fn test_when_used_with_assignment_and_expression_result() {
    let out = run_prints(r#"
        fun build(value: Int): String {
            val out = when (value) {
                1 -> "one"
                2 -> "two"
                else -> "other"
            }
            return out
        }

        fun main() {
            println(build(1))
            println(build(2))
            println(build(9))
        }
    "#);
    assert_eq!(out, &["one", "two", "other"]);
}

#[test]
fn test_when_with_fallback_expressions_using_arithmetic() {
    let out = run_prints(r#"
        fun score(value: Int): String {
            return when (value % 3) {
                0 -> "triple"
                1 -> "plus"
                2 -> "plus2"
                else -> "?"
            }
        }

        fun main() {
            println(score(10))
            println(score(11))
            println(score(12))
        }
    "#);
    assert_eq!(out, &["plus", "plus2", "triple"]);
}

#[test]
fn test_when_with_local_type_checks_and_smart_casts() {
    let out = run_prints(r#"
        fun convert(value: Any): String {
            return when (value) {
                is Int -> "i=" + value.toString()
                is Long -> "l=" + value.toString()
                is Double -> "d=" + value.toString()
                else -> "x"
            }
        }

        fun main() {
            println(convert(3))
            println(convert(4L))
            println(convert(1.5))
            println(convert("x"))
        }
    "#);
    assert_eq!(out, &["i=3", "l=4", "d=1.5", "x"]);
}

#[test]
fn test_when_on_data_class_property_subject() {
    let out = run_prints(r#"
        data class User(val name: String, val active: Boolean, val level: Int)

        fun label(user: User): String {
            return when {
                user.name.isEmpty() -> "anon"
                !user.active -> "inactive"
                user.level > 10 -> "vip"
                else -> "regular"
            }
        }

        fun main() {
            println(label(User("", true, 3)))
            println(label(User("a", false, 1)))
            println(label(User("b", true, 12)))
            println(label(User("c", true, 4)))
        }
    "#);
    assert_eq!(out, &["anon", "inactive", "vip", "regular"]);
}

#[test]
fn test_when_with_nested_subject_binding_in_same_when() {
    let out = run_prints(r#"
        fun decode(value: Any): String {
            return when (value) {
                is Int -> {
                    val doubled = value * 2
                    when {
                        doubled > 10 -> "int-big"
                        else -> "int-small"
                    }
                }
                is String -> {
                    val head = value.firstOrNull() ?: '?'
                    when (head) {
                        in 'a'..'m' -> "string-low"
                        in 'n'..'z' -> "string-high"
                        else -> "string-other"
                    }
                }
                else -> "none"
            }
        }

        fun main() {
            println(decode(7))
            println(decode(4))
            println(decode("beta"))
            println(decode("zeta"))
            println(decode("@"))
            println(decode(3.0))
        }
    "#);
    assert_eq!(out, &["int-small", "int-small", "string-low", "string-high", "string-other", "none"]);
}

#[test]
fn test_when_with_throws_in_subject_branches() {
    let out = run_prints(r#"
        fun classify(n: Int): String {
            return when (n) {
                0 -> "zero"
                1 -> "one"
                in 2..9 -> "few"
                else -> throw Error("too-large")
            }
        }

        fun main() {
            try {
                println(classify(1))
                println(classify(5))
                println(classify(20))
            } catch (e: Error) {
                println("error")
            }
        }
    "#);
    assert_eq!(out, &["one", "few", "error"]);
}

#[test]
fn test_when_with_sealed_like_polymorphic_input() {
    let out = run_prints(r#"
        sealed class Node {
            class A(val value: Int) : Node()
            class B(val value: String) : Node()
            class C : Node()
        }

        fun render(node: Node): String {
            return when (node) {
                is Node.A -> "A:" + node.value
                is Node.B -> "B:" + node.value
                is Node.C -> "C"
            }
        }

        fun main() {
            println(render(Node.A(9)))
            println(render(Node.B("x")))
            println(render(Node.C()))
        }
    "#);
    assert_eq!(out, &["A:9", "B:x", "C"]);
}

#[test]
fn test_when_statement_without_else_on_exhaustive_input() {
    let out = run_prints(r#"
        fun colorCode(color: String): Int {
            return when (color) {
                "red" -> 1
                "green" -> 2
                "blue" -> 3
                else -> 0
            }
        }

        fun main() {
            println(colorCode("red"))
            println(colorCode("green"))
            println(colorCode("blue"))
            println(colorCode("black"))
        }
    "#);
    assert_eq!(out, &["1", "2", "3", "0"]);
}

#[test]
fn test_when_with_subject_in_function_reference_style() {
    let out = run_prints(r#"
        fun classify(value: Int): String {
            val fn = { n: Int ->
                when (n) {
                    1 -> "single"
                    2, 3 -> "pair"
                    in 4..6 -> "few"
                    else -> "many"
                }
            }
            return fn(value)
        }

        fun main() {
            println(classify(1))
            println(classify(3))
            println(classify(5))
            println(classify(9))
        }
    "#);
    assert_eq!(out, &["single", "pair", "few", "many"]);
}

#[test]
fn test_when_with_subject_value_block_scope() {
    let out = run_prints(r#"
        fun render(level: Int): String {
            return when (level) {
                in 0..9 -> {
                    val label = "low"
                    label + ":" + level
                }
                in 10..19 -> {
                    val offset = level - 10
                    "mid:" + offset
                }
                else -> {
                    val doubled = level * 2
                    "high:" + doubled
                }
            }
        }

        fun main() {
            println(render(4))
            println(render(13))
            println(render(30))
        }
    "#);
    assert_eq!(out, &["low:4", "mid:3", "high:60"]);
}

#[test]
fn test_when_expression_with_guarded_ranges_and_type_checks() {
    let out = run_prints(r#"
        fun describe(value: Any): String {
            return when {
                value is Int && value < 0 -> "negative"
                value is Int && value == 0 -> "zero"
                value is Int -> "positive"
                value is String && value.isNotBlank() -> "word"
                value == null -> "none"
                else -> "other"
            }
        }

        fun main() {
            println(describe(-1))
            println(describe(0))
            println(describe(5))
            println(describe("a"))
            println(describe(null))
        }
    "#);
    assert_eq!(out, &["negative", "zero", "positive", "word", "none"]);
}
