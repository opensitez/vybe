use crate::helpers::run_prints;

#[test]
fn test_hello_world() {
    let out = run_prints(
        r#"
        fun main() {
            println("Hello, Kotlin!")
        }
    "#,
    );
    assert_eq!(out, &["Hello, Kotlin!"]);
}

#[test]
fn test_variables_and_arithmetic() {
    let out = run_prints(
        r#"
        fun main() {
            val a = 15
            var b = 25
            val c = a + b
            println(c)
        }
    "#,
    );
    assert_eq!(out, &["40"]);
}

#[test]
fn test_string_template_var() {
    let out = run_prints(
        r#"
        fun main() {
            val name = "Vybe"
            println("Hello $name")
        }
    "#,
    );
    assert_eq!(out, &["Hello Vybe"]);
}

#[test]
fn test_val_declaration() {
    let out = run_prints(
        r#"
        fun main() {
            val x = 100
            println(x)
        }
    "#,
    );
    assert_eq!(out, &["100"]);
}

#[test]
fn test_var_reassignment() {
    let out = run_prints(
        r#"
        fun main() {
            var x = 10
            x = 20
            println(x)
        }
    "#,
    );
    assert_eq!(out, &["20"]);
}

#[test]
fn test_addition_subtraction() {
    let out = run_prints(
        r#"
        fun main() {
            val a = 50
            val b = 20
            println(a + b)
            println(a - b)
        }
    "#,
    );
    assert_eq!(out, &["70", "30"]);
}

#[test]
fn test_multiplication_division() {
    let out = run_prints(
        r#"
        fun main() {
            val a = 12
            val b = 4
            println(a * b)
            println(a / b)
        }
    "#,
    );
    assert_eq!(out, &["48", "3"]);
}

#[test]
fn test_modulo_operator() {
    let out = run_prints(
        r#"
        fun main() {
            val a = 17
            val b = 5
            println(a % b)
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_compound_add_assign() {
    let out = run_prints(
        r#"
        fun main() {
            var x = 10
            x += 5
            println(x)
        }
    "#,
    );
    assert_eq!(out, &["15"]);
}

#[test]
fn test_compound_sub_assign() {
    let out = run_prints(
        r#"
        fun main() {
            var x = 20
            x -= 8
            println(x)
        }
    "#,
    );
    assert_eq!(out, &["12"]);
}

#[test]
fn test_compound_mul_assign() {
    let out = run_prints(
        r#"
        fun main() {
            var x = 6
            x *= 7
            println(x)
        }
    "#,
    );
    assert_eq!(out, &["42"]);
}

#[test]
fn test_compound_div_assign() {
    let out = run_prints(
        r#"
        fun main() {
            var x = 100
            x /= 4
            println(x)
        }
    "#,
    );
    assert_eq!(out, &["25"]);
}

#[test]
fn test_compound_mod_assign() {
    let out = run_prints(
        r#"
        fun main() {
            var x = 29
            x %= 6
            println(x)
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_operator_precedence_1() {
    let out = run_prints(
        r#"
        fun main() {
            val res = 2 + 3 * 4
            println(res)
        }
    "#,
    );
    assert_eq!(out, &["14"]);
}

#[test]
fn test_operator_precedence_2() {
    let out = run_prints(
        r#"
        fun main() {
            val res = (2 + 3) * 4
            println(res)
        }
    "#,
    );
    assert_eq!(out, &["20"]);
}

#[test]
fn test_boolean_and() {
    let out = run_prints(
        r#"
        fun main() {
            println(true && true)
            println(true && false)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_boolean_or() {
    let out = run_prints(
        r#"
        fun main() {
            println(true || false)
            println(false || false)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_boolean_not() {
    let out = run_prints(
        r#"
        fun main() {
            println(!true)
            println(!false)
        }
    "#,
    );
    assert_eq!(out, &["false", "true"]);
}

#[test]
fn test_boolean_complex() {
    let out = run_prints(
        r#"
        fun main() {
            val a = true
            val b = false
            val c = true
            println((a || b) && c)
        }
    "#,
    );
    assert_eq!(out, &["true"]);
}

#[test]
fn test_integer_literals() {
    let out = run_prints(
        r#"
        fun main() {
            val a = 0
            val b = 42
            println(a)
            println(b)
        }
    "#,
    );
    assert_eq!(out, &["0", "42"]);
}

#[test]
fn test_float_literals() {
    let out = run_prints(
        r#"
        fun main() {
            val f = 3.14
            println(f)
        }
    "#,
    );
    assert_eq!(out, &["3.14"]);
}

#[test]
fn test_comparison_equal() {
    let out = run_prints(
        r#"
        fun main() {
            println(10 == 10)
            println(10 == 20)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_string_concatenation() {
    let out = run_prints(
        r#"
        fun main() {
            val first = "Hello"
            val second = "Kotlin"
            println(first + ", " + second)
        }
    "#,
    );
    assert_eq!(out, &["Hello, Kotlin"]);
}

#[test]
fn test_string_template() {
    let out = run_prints(
        r#"
        fun main() {
            val name = "Kotlin"
            val version = 1
            println("Language ${name} ${version + 1}")
        }
    "#,
    );
    assert_eq!(out, &["Language Kotlin 2"]);
}

#[test]
fn test_float_division() {
    let out = run_prints(
        r#"
        fun main() {
            val result = 7 / 2.0
            println(result)
        }
    "#,
    );
    assert_eq!(out, &["3.5"]);
}

#[test]
fn test_comparison_chain() {
    let out = run_prints(
        r#"
        fun main() {
            val score = 82
            println(score > 80 && score < 90)
            println(score < 50 || score == 82)
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_prefix_increment() {
    let out = run_prints(
        r#"
        fun main() {
            var count = 1
            println(++count)
            println(--count)
        }
    "#,
    );
    assert_eq!(out, &["2", "1"]);
}

#[test]
fn test_unary_negation_and_plus() {
    let out = run_prints(
        r#"
        fun main() {
            val positive = +5
            val negative = -positive
            println(negative)
        }
    "#,
    );
    assert_eq!(out, &["-5"]);
}

#[test]
fn test_equality_and_inequality() {
    let out = run_prints(
        r#"
        fun main() {
            println("a" == "a")
            println("a" != "b")
            println(10 <= 10)
            println(9 >= 10)
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true", "false"]);
}

#[test]
fn test_comparison_not_equal() {
    let out = run_prints(
        r#"
        fun main() {
            println(10 != 20)
            println(10 != 10)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_comparison_less_greater() {
    let out = run_prints(
        r#"
        fun main() {
            println(5 < 10)
            println(15 > 10)
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_comparison_less_equal_greater_equal() {
    let out = run_prints(
        r#"
        fun main() {
            println(10 <= 10)
            println(10 >= 10)
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_string_concatenation_simple_join() {
    let out = run_prints(
        r#"
fun main() {
val s = "Hello " + "World"
            println(s)
        }
    "#,
    );
    assert_eq!(out, &["Hello World"]);
}

#[test]
fn test_string_template_multiple_vars() {
    let out = run_prints(
        r#"
        fun main() {
            val lang = "Kotlin"
            val ver = "1.9"
            println("$lang $ver")
        }
    "#,
    );
    assert_eq!(out, &["Kotlin 1.9"]);
}

#[test]
fn test_nested_block_scopes() {
    // A BARE `{}` statement in Kotlin is an unexecuted lambda literal —
    // `run { }` is the valid spelling of a scoped block.
    let out = run_prints(
        r#"
        fun main() {
            val x = 1
            run {
                val y = 2
                println(x + y)
            }
            println(x)
        }
    "#,
    );
    assert_eq!(out, &["3", "1"]);
}

#[test]
fn test_grouped_expressions() {
    let out = run_prints(
        r#"
        fun main() {
            val a = (10 + 20) * (30 - 10)
            println(a)
        }
    "#,
    );
    assert_eq!(out, &["600"]);
}

#[test]
fn test_if_expression_used_as_return_value() {
    let out = run_prints(
        r#"
        fun classify(value: Int): String {
            return if (value > 10) "big" else "small"
        }

        fun main() {
            println(classify(11))
            println(classify(3))
        }
    "#,
    );
    assert_eq!(out, &["big", "small"]);
}

#[test]
fn test_char_literals_and_conversion() {
    let out = run_prints(
        r#"
        fun main() {
            val first: Char = 'A'
            val second = 'b'
            println(first)
            println(second)
            println(first.code + 1)
            println(second.code)
        }
    "#,
    );
    assert_eq!(out, &["A", "b", "66", "98"]);
}

#[test]
fn test_nullable_reference_in_basic_flow() {
    let out = run_prints(
        r#"
        fun main() {
            val maybe: String? = null
            val value: String = maybe ?: "fallback"
            println(value)
            val explicit: String? = "ok"
            println(explicit ?: "fallback")
        }
    "#,
    );
    assert_eq!(out, &["fallback", "ok"]);
}

#[test]
fn test_block_scoped_shadowing_and_assignment() {
    let out = run_prints(
        r#"
        fun main() {
            val x = 10
            run {
                val x = 99
                println(x)
            }
            println(x)
        }
    "#,
    );
    assert_eq!(out, &["99", "10"]);
}

#[test]
fn test_raw_multiline_string_preserves_content() {
    let out = run_prints(
        r#"
        fun main() {
            val text = """
first
second
third
"""
            println(text.trim())
        }
    "#,
    );
    assert_eq!(out, &["first\nsecond\nthird"]);
}

#[test]
fn test_string_escape_sequences() {
    let out = run_prints(
        r#"
        fun main() {
            println("a\\nb\\n")
            println("tab\\tend")
            println("quote: \"")
            println('c')
        }
    "#,
    );
    assert_eq!(out, &["a\\nb\\n", "tab\\tend", "quote: \"", "c"]);
}

#[test]
fn test_boolean_operator_short_circuit_false_guard() {
    let out = run_prints(
        r#"
        var called = 0

        fun sideEffect(): Boolean {
            called += 1
            return true
        }

        fun main() {
            println(false && sideEffect())
            println(called)
            println(true && sideEffect())
            println(called)
        }
    "#,
    );
    assert_eq!(out, &["false", "0", "true", "1"]);
}

#[test]
fn test_boolean_operator_short_circuit_true_or() {
    let out = run_prints(
        r#"
        var called = 0

        fun sideEffect(): Boolean {
            called += 1
            return false
        }

        fun main() {
            println(true || sideEffect())
            println(called)
            println(false || sideEffect())
            println(called)
        }
    "#,
    );
    assert_eq!(out, &["true", "0", "false", "1"]);
}

#[test]
fn test_modulo_sign_edges() {
    let out = run_prints(
        r#"
        fun main() {
            println((-10) % 3)
            println(10 % (-3))
            println((-10) % (-3))
        }
    "#,
    );
    // Kotlin's % follows the DIVIDEND's sign: 10 % -3 is 1 (real Kotlin
    // agrees).
    assert_eq!(out, &["-1", "1", "-1"]);
}

#[test]
fn test_parenthesized_condition_precedence_in_if() {
    let out = run_prints(
        r#"
        fun main() {
            val result = (true || false) && false
            println(result)
            val result2 = true || (false && false)
            println(result2)
        }
    "#,
    );
    assert_eq!(out, &["false", "true"]);
}
