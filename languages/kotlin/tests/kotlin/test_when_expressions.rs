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
