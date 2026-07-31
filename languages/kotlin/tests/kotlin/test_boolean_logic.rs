use crate::helpers::run_prints;

#[test]
fn test_boolean_literals_and_basic_ops() {
    let out = run_prints(r#"
        fun main() {
            println(true)
            println(false)
            println(true && false)
            println(true || false)
            println(!true)
            println(!false)
        }
    "#);
    assert_eq!(out, &["true", "false", "false", "true", "false", "true"]);
}

#[test]
fn test_boolean_or_and_precedence() {
    let out = run_prints(r#"
        fun main() {
            println(true || false && false)
            println((true || false) && false)
            println(false || true && true)
            println((false || true) && true)
        }
    "#);
    assert_eq!(out, &["true", "false", "true", "true"]);
}

#[test]
fn test_boolean_xor_behavior() {
    let out = run_prints(r#"
        fun main() {
            println(true xor true)
            println(true xor false)
            println(false xor true)
            println(false xor false)
        }
    "#);
    assert_eq!(out, &["false", "true", "true", "false"]);
}

#[test]
fn test_boolean_short_circuit_and_side_effects() {
    let out = run_prints(r#"
        fun main() {
            var calls = 0
            fun shouldRun() : Boolean {
                calls++
                return true
            }
            println(false && shouldRun())
            println(calls)
            println(true || shouldRun())
            println(calls)
        }
    "#);
    assert_eq!(out, &["false", "0", "true", "0"]);
}

#[test]
fn test_boolean_short_circuit_with_failure_paths() {
    let out = run_prints(r#"
        fun main() {
            var calls = 0
            fun sideEffect() : Boolean {
                calls++
                throw Error("boom")
            }
            println(false && sideEffect())
            println(calls)
            println(true || sideEffect())
            println(calls)
        }
    "#);
    assert_eq!(out, &["false", "0", "true", "0"]);
}

#[test]
fn test_boolean_eager_evaluation_of_if_conditions() {
    let out = run_prints(r#"
        fun main() {
            val a = true
            val b = false
            val c = if (a && b) "both" else "not both"
            val d = if (a || b) "some" else "none"
            println(c)
            println(d)
        }
    "#);
    assert_eq!(out, &["not both", "some"]);
}

#[test]
fn test_boolean_in_when_conditions() {
    let out = run_prints(r#"
        fun main() {
            val value = 2
            val result = when {
                value % 2 == 0 && value > 1 -> "even"
                value < 0 && value % 2 == 0 -> "neg"
                else -> "other"
            }
            println(result)
        }
    "#);
    assert_eq!(out, &["even"]);
}

#[test]
fn test_boolean_for_range_filters_and_predicates() {
    let out = run_prints(r#"
        fun main() {
            val nums = intArrayOf(1, 2, 3, 4, 5)
            val evens = nums.filter { it % 2 == 0 }
            val odds = nums.filter { it % 2 != 0 }
            println(evens.joinToString(","))
            println(odds.joinToString(","))
        }
    "#);
    assert_eq!(out, &["2,4", "1,3,5"]);
}

#[test]
fn test_boolean_equality_and_identity() {
    let out = run_prints(r#"
        fun main() {
            println(true == true)
            println(true == false)
            println(false == false)
            val a: Boolean = false
            val b = a
            println(a === b)
        }
    "#);
    assert_eq!(out, &["true", "false", "true", "true"]);
}

#[test]
fn test_boolean_to_string_output() {
    let out = run_prints(r#"
        fun main() {
            println(true.toString())
            println(false.toString())
            println((true && true).toString())
        }
    "#);
    assert_eq!(out, &["true", "false", "true"]);
}

#[test]
fn test_boolean_from_string_conversion() {
    let out = run_prints(r#"
        fun main() {
            println("true".toBoolean())
            println("false".toBoolean())
            println("TRUE".toBoolean())
            println("junk".toBoolean())
        }
    "#);
    assert_eq!(out, &["true", "false", "false", "false"]);
}

#[test]
fn test_boolean_nullable_coercion() {
    let out = run_prints(r#"
        fun main() {
            val value: Boolean? = null
            val fallback = value ?: false
            println(fallback)
            val value2: Boolean? = true
            println(value2 ?: false)
        }
    "#);
    assert_eq!(out, &["false", "true"]);
}

#[test]
fn test_boolean_safe_call_and_elvis() {
    let out = run_prints(r#"
        fun main() {
            val value: Boolean? = null
            println(value?.toString() ?: "missing")
            val another: Boolean? = false
            println(another?.toString() ?: "missing")
        }
    "#);
    assert_eq!(out, &["missing", "false"]);
}

#[test]
fn test_boolean_let_and_scope_functions() {
    let out = run_prints(r#"
        fun main() {
            val value: Boolean? = true
            val transformed = value?.let { it && true } ?: false
            println(transformed)
            val none: Boolean? = null
            println(none?.let { it } ?: false)
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_boolean_and_comparison_chains() {
    let out = run_prints(r#"
        fun main() {
            val a = 3
            val b = 4
            println(a < b && a == 3)
            println(a > b || b == 4)
            println(a <= b && b % 2 == 0)
            println(a in 1..b)
        }
    "#);
    assert_eq!(out, &["true", "true", "false", "true"]);
}

#[test]
fn test_boolean_operator_normalization_in_loops() {
    let out = run_prints(r#"
        fun main() {
            var i = 0
            var ok = true
            while (ok && i < 3) {
                i++
                if (i == 2) {
                    ok = false
                }
            }
            println(i)
            println(ok)
        }
    "#);
    assert_eq!(out, &["2", "false"]);
}

#[test]
fn test_boolean_negation_chain() {
    let out = run_prints(r#"
        fun main() {
            val a = true
            val b = false
            println(!a)
            println(!!a)
            println(!!!b)
            println(!(!a && b))
        }
    "#);
    assert_eq!(out, &["false", "true", "true", "true"]);
}

#[test]
fn test_boolean_reified_conditional_outputs() {
    let out = run_prints(r#"
        fun main() {
            fun status(flag: Boolean): String {
                return if (!flag) "off" else "on"
            }
            println(status(false))
            println(status(true))
            println(status(status(false) == "off"))
        }
    "#);
    assert_eq!(out, &["off", "on", "on"]);
}

#[test]
fn test_boolean_array_contains_all() {
    let out = run_prints(r#"
        fun main() {
            val values = booleanArrayOf(true, false, true)
            println(values.all { it })
            println(values.any { it })
            println(values.count { it })
            println(values.count { !it })
        }
    "#);
    assert_eq!(out, &["false", "true", "2", "1"]);
}

#[test]
fn test_boolean_fold_reductions() {
    let out = run_prints(r#"
        fun main() {
            val values = booleanArrayOf(true, false, true, true)
            println(values.fold(true) { acc, value -> acc && value })
            println(values.fold(false) { acc, value -> acc || value })
            println(values.reduce { acc, value -> acc && value })
            println(values.reduce { acc, value -> acc || value })
        }
    "#);
    assert_eq!(out, &["false", "true", "false", "true"]);
}

#[test]
fn test_boolean_join_to_string_and_join() {
    let out = run_prints(r#"
        fun main() {
            val flags = booleanArrayOf(true, false, true)
            val text = flags.joinToString("|") { it.toString() }
            println(text)
            println(text.length)
        }
    "#);
    assert_eq!(out, &["true|false|true", "14"]);
}

#[test]
fn test_boolean_ordering_via_to_int_like_comparison() {
    let out = run_prints(r#"
        fun main() {
            val a = true
            val b = false
            println(if (a == b) 0 else if (a && !b) 1 else -1)
            println(if (!a == b) "swap" else "noswap")
        }
    "#);
    assert_eq!(out, &["1", "noswap"]);
}

#[test]
fn test_boolean_with_while_break_continue_controls() {
    let out = run_prints(r#"
        fun main() {
            var i = 0
            var total = 0
            while (i < 6) {
                i++
                if (i % 2 == 0) {
                    continue
                }
                if (i > 4) {
                    break
                }
                total += i
            }
            println(i)
            println(total)
        }
    "#);
    assert_eq!(out, &["5", "4"]);
}

#[test]
fn test_boolean_guarded_mapping() {
    let out = run_prints(r#"
        fun main() {
            val input = listOf(1, 2, 3, 4, 5)
            val filtered = input.filter { it > 2 && it < 5 }
            val mapped = filtered.map { it % 2 == 0 }
            println(filtered.joinToString(","))
            println(mapped.joinToString(","))
            println(filtered.size)
        }
    "#);
    assert_eq!(out, &["3,4", "false,true", "2"]);
}

#[test]
fn test_boolean_short_circuit_in_while_conditions() {
    let out = run_prints(r#"
        fun main() {
            var i = 0
            var steps = 0
            while (i < 3 && steps < 5) {
                steps++
                i++
            }
            println(i)
            println(steps)
            println(i == 3 && steps == 3)
        }
    "#);
    assert_eq!(out, &["3", "3", "true"]);
}

#[test]
fn test_boolean_arithmetic_operators_and_assignment_mix() {
    let out = run_prints(r#"
        fun main() {
            var a = true
            var b = false
            a = a && !b
            b = a || !b
            println(a)
            println(b)
            a = a.xor(b)
            println(a)
        }
    "#);
    assert_eq!(out, &["true", "true", "false"]);
}

#[test]
fn test_boolean_if_else_chain() {
    let out = run_prints(r#"
        fun main() {
            val value = 7
            val status = if (value > 10) false
                         else if (value > 5 && value < 10) true
                         else false
            println(status)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_boolean_expression_in_for_filter() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf("a", "bb", "ccc")
            val ok = values.filter { it.length > 1 && it.length < 3 }
            println(ok.joinToString(","))
            val fail = values.filter { !ok.contains(it) }
            println(fail.joinToString(","))
        }
    "#);
    assert_eq!(out, &["bb", "a,ccc"]);
}

#[test]
fn test_boolean_negated_predicates_for_classification() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(-2, -1, 0, 1, 2)
            val positives = values.filter { it > 0 }
            val nonPositive = values.filter { !(it > 0) }
            println(positives.joinToString(","))
            println(nonPositive.joinToString(","))
            println(nonPositive.all { ! (it > 0) })
        }
    "#);
    assert_eq!(out, &["1,2", "-2,-1,0", "true"]);
}

#[test]
fn test_boolean_mix_with_pattern_match_like_when() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(true, false, true)
            var trues = 0
            var falses = 0
            for (value in values) {
                when (value) {
                    true -> trues++
                    false -> falses++
                }
            }
            println(trues)
            println(falses)
        }
    "#);
    assert_eq!(out, &["2", "1"]);
}
