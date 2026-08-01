use crate::helpers::run_prints;

#[test]
fn test_function_args_evaluate_left_to_right() {
    let out = run_prints(
        r#"
        fun main() {
            var order = ""
            fun lhs(): Int { order += "L"; return 1 }
            fun mid(v: Int): Int { order += "M"; return v + 1 }
            fun rhs(v: Int): Int { order += "R"; return v + 2 }
            fun f(a: Int, b: Int, c: Int) {
                println(a)
                println(b)
                println(c)
            }
            f(lhs(), mid(0), rhs(0))
            println(order)
        }
    "#,
    );
    assert_eq!(out, &["1", "1", "2", "LMR"]);
}

#[test]
fn test_default_arguments_evaluate_when_omitted() {
    let out = run_prints(
        r#"
        fun main() {
            var order = ""
            fun value(prefix: String = "d"): String {
                order += "v"
                return prefix
            }
            fun report(a: String = value("a"), b: String = value("b")) {
                println(a)
                println(b)
            }
            report("x")
            println(order)
        }
    "#,
    );
    assert_eq!(out, &["x", "b", "v"]);
}

#[test]
fn test_constructor_arg_order() {
    let out = run_prints(
        r#"
        class Tracker {
            constructor(a: Int, b: Int) {
                println(a)
                println(b)
            }
        }

        fun main() {
            var log = ""
            fun first(): Int { log += "1"; return 1 }
            fun second(): Int { log += "2"; return 2 }
            Tracker(first(), second())
            println(log)
        }
    "#,
    );
    assert_eq!(out, &["1", "2", "12"]);
}

#[test]
fn test_property_initializer_runs_before_next_access() {
    let out = run_prints(
        r#"
        var order = ""
        val one = run {
            order += "1"
            1
        }
        val two = run {
            order += "2"
            2
        }
        fun main() {
            println(one + two)
            println(order)
        }
    "#,
    );
    assert_eq!(out, &["3", "12"]);
}

#[test]
fn test_when_subject_and_guard_order() {
    let out = run_prints(
        r#"
        fun main() {
            var order = ""
            val v = run {
                order += "s"
                7
            }
            val out = when {
                (order += "g").isNotEmpty() && v > 0 -> {
                    order += "t"
                    "yes"
                }
                else -> {
                    order += "f"
                    "no"
                }
            }
            println(out)
            println(order)
        }
    "#,
    );
    assert_eq!(out, &["yes", "sgt"]);
}

#[test]
fn test_elvis_uses_rhs_only_when_lhs_null() {
    let out = run_prints(
        r#"
        fun main() {
            val lhs: String? = null
            var order = ""
            val rhs = run { order += "rhs"; "value" }
            val out = lhs ?: rhs
            println(out)
            println(order)
        }
    "#,
    );
    assert_eq!(out, &["value", "rhs"]);
}

#[test]
fn test_elvis_skips_rhs_when_non_null() {
    let out = run_prints(
        r#"
        fun main() {
            val lhs: String? = "left"
            var order = ""
            val out = lhs ?: run { order += "rhs"; "value" }
            println(out)
            println(order)
        }
    "#,
    );
    assert_eq!(out, &["left", ""]);
}

#[test]
fn test_range_loop_evaluates_bounds_left_to_right() {
    let out = run_prints(
        r#"
        fun main() {
            var order = ""
            val start = run { order += "s"; 1 }
            val end = run { order += "e"; 3 }
            val values = (start..end).toList().joinToString(",")
            println(values)
            println(order)
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "se"]);
}

#[test]
fn test_binary_ops_left_to_right_operands() {
    let out = run_prints(
        r#"
        fun main() {
            var order = ""
            fun left(): Int { order += "L"; return 1 }
            fun right(): Int { order += "R"; return 2 }
            println(left() + right())
            println(order)
        }
    "#,
    );
    assert_eq!(out, &["3", "LR"]);
}

#[test]
fn test_assignment_expression_order() {
    let out = run_prints(
        r#"
        fun main() {
            var a = 0
            var b = 0
            val list = mutableListOf<Int>()
            fun left() { a = 1; list.add(1) }
            fun right() { b = 2; list.add(2) }
            left()
            right()
            println(a + b)
            println(list.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["3", "1,2"]);
}

#[test]
fn test_nested_calls_evaluate_outer_before_inner_return() {
    let out = run_prints(
        r#"
        fun combine(x: Int, y: Int): Int = x * 10 + y
        fun main() {
            var order = ""
            fun f(): Int { order += "f"; return 1 }
            fun g(x: Int): Int { order += "g"; return x + 1 }
            fun h(x: Int): Int { order += "h"; return x + 2 }
            val out = combine(g(f()), h(g(3)))
            println(out)
            println(order)
        }
    "#,
    );
    assert_eq!(out, &["26", "fghg"]);
}

#[test]
fn test_for_each_loop_evaluates_initializer_before_body() {
    let out = run_prints(
        r#"
        fun main() {
            var order = ""
            val list = run {
                order += "init"
                listOf(1, 2)
            }
            var sum = 0
            for (x in list) {
                order += "-"
                sum += x
            }
            println(sum)
            println(order)
        }
    "#,
    );
    assert_eq!(out, &["3", "init--"]);
}

#[test]
fn test_ternary_like_when_not_available() {
    let out = run_prints(
        r#"
        fun main() {
            val value = if (true) "yes" else "no"
            println(value)
        }
    "#,
    );
    assert_eq!(out, &["yes"]);
}

#[test]
fn test_conditional_operator_right_hand_only_on_false() {
    let out = run_prints(
        r#"
        fun main() {
            var log = ""
            val value = if (true) { log += "t"; 1 } else { log += "f"; 2 }
            println(value)
            println(log)
        }
    "#,
    );
    assert_eq!(out, &["1", "t"]);
}

#[test]
fn test_early_return_in_lambda_skips_subsequent_calls() {
    let out = run_prints(
        r#"
        fun main() {
            var order = ""
            fun side(): Int {
                order += "s"
                return 2
            }
            val out = run {
                order += "r"
                1
            }
            if (out > 0) {
                println(side())
            } else {
                println(0)
            }
            println(order)
        }
    "#,
    );
    assert_eq!(out, &["2", "rs"]);
}

#[test]
fn test_not_operator_takes_single_evaluation() {
    let out = run_prints(
        r#"
        fun main() {
            var order = ""
            fun flag(): Boolean {
                order += "f"
                return false
            }
            println(!flag())
            println(order)
        }
    "#,
    );
    assert_eq!(out, &["true", "f"]);
}

#[test]
fn test_multiple_expressions_in_block_preserve_order() {
    let out = run_prints(
        r#"
        fun main() {
            val order = run {
                val out = StringBuilder()
                out.append("a")
                out.append("b")
                out.append("c")
                out.toString()
            }
            println(order)
        }
    "#,
    );
    assert_eq!(out, &["abc"]);
}

#[test]
fn test_list_of_with_mixed_evals_in_initializer() {
    let out = run_prints(
        r#"
        fun main() {
            var order = ""
            fun a(): Int { order += "A"; return 1 }
            fun b(): Int { order += "B"; return 2 }
            fun c(): Int { order += "C"; return 3 }
            val values = listOf(a(), b(), c())
            println(values.joinToString(","))
            println(order)
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "ABC"]);
}

#[test]
fn test_string_interpolation_eval_order_left_to_right() {
    let out = run_prints(
        r#"
        fun main() {
            var order = ""
            fun left(): String { order += "L"; return "x" }
            fun right(): String { order += "R"; return "y" }
            val result = "${left()}${right()}"
            println(result)
            println(order)
        }
    "#,
    );
    assert_eq!(out, &["xy", "LR"]);
}

#[test]
fn test_operator_precedence_evaluation_still_left_to_right_for_same_level() {
    let out = run_prints(
        r#"
        fun main() {
            var order = ""
            fun a(): Int { order += "a"; return 1 }
            fun b(): Int { order += "b"; return 2 }
            fun c(): Int { order += "c"; return 3 }
            val out = a() + b() * c()
            println(out)
            println(order)
        }
    "#,
    );
    assert_eq!(out, &["7", "abc"]);
}

#[test]
fn test_method_chain_evaluates_receiver_then_argument() {
    let out = run_prints(
        r#"
        fun main() {
            var order = ""
            fun arg(v: Int): Int {
                order += "a" + v
                return v
            }
            val out = listOf(1, 2, 3)
                .map { arg(it) }
                .sum()
            println(out)
            println(order)
        }
    "#,
    );
    assert_eq!(out, &["6", "a1a2a3"]);
}
