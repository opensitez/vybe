kotlin_run_test!(
    test_do_while_executes_once,
    r#"
        fun main() {
            var i = 10
            var seen = 0
            do {
                seen += 1
            } while (i < 0)
            println(seen)
        }
    "#,
    &["1"]
);

kotlin_run_test!(
    test_do_while_accumulates_while_condition,
    r#"
        fun main() {
            var i = 0
            var out = 0
            do {
                out += i
                i += 1
            } while (i < 4)
            println(out)
            println(i)
        }
    "#,
    &["6", "4"]
);

kotlin_run_test!(
    test_do_while_with_break,
    r#"
        fun main() {
            var i = 0
            var out = 0
            do {
                if (i == 2) break
                out += i
                i += 1
            } while (true)
            println(out)
            println(i)
        }
    "#,
    &["1", "0"]
);

kotlin_run_test!(
    test_do_while_continue_shape,
    r#"
        fun main() {
            var i = 0
            var out = 0
            do {
                i += 1
                if (i == 2) continue
                out += i
            } while (i < 5)
            println(out)
        }
    "#,
    &["13"]
);

kotlin_run_test!(
    test_do_while_nested,
    r#"
        fun main() {
            var outer = 0
            var inner = 0
            do {
                outer += 1
                var j = 0
                do {
                    inner += j
                    j += 1
                } while (j < 2)
            } while (outer < 3)
            println(outer)
            println(inner)
        }
    "#,
    &["3", "3"]
);

kotlin_run_test!(
    test_do_while_false_predicate_after_increment,
    r#"
        fun main() {
            var i = 0
            var out = 0
            do {
                out += i
                i += 2
            } while (i > 10)
            println(out)
        }
    "#,
    &["0"]
);

kotlin_run_test!(
    test_do_while_negative_start,
        r#"
        fun main() {
            var i = -3
            var out = 0
            do {
                out += i
                i += 1
            } while (i <= 0)
            println(out)
            println(i)
        }
    "#,
    &["-6", "1"]
);

kotlin_run_test!(
    test_do_while_function_guard,
    r#"
        fun keep(i: Int): Boolean = i < 3
        fun main() {
            var i = 0
            var out = ""
            do {
                out += i.toString()
                i += 1
            } while (keep(i))
            println(out)
        }
    "#,
    &["012"]
);

kotlin_run_test!(
    test_do_while_boolean_expression,
    r#"
        fun main() {
            var i = 0
            var out = 0
            do {
                out += if (i % 2 == 0) i else 0
                i += 1
            } while (i < 6)
            println(out)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_do_while_with_local_scope,
    r#"
        fun main() {
            var total = 0
            do {
                val step = 2
                total += step
            } while (total < 8)
            println(total)
        }
    "#,
    &["10"]
);

kotlin_run_test!(
        test_do_while_expression_result,
    r#"
        fun main() {
            var x = 0
            val value = kotlin.run {
                x = 1
                x
            }
            println(x)
            println(value)
        }
    "#,
    &["1", "1"]
);

kotlin_run_test!(
    test_do_while_with_continue_after_limit,
    r#"
        fun main() {
            var i = 0
            var total = 0
            do {
                i += 1
                if (i == 3) continue
                total += i
            } while (i < 6)
            println(total)
        }
    "#,
    &["16"]
);

kotlin_run_test!(
    test_do_while_complex_condition,
    r#"
        fun shouldRun(x: Int): Boolean = x <= 2
        fun main() {
            var i = 0
            var total = 0
            do {
                total += i
                i += 1
            } while (i < 6 && shouldRun(i))
            println(total)
            println(i)
        }
    "#,
    &["3", "3"]
);

kotlin_run_test!(
    test_do_while_string_builder,
    r#"
        fun main() {
            var i = 0
            var out = ""
            do {
                out += i.toString()
                i += 1
            } while (i < 4)
            println(out)
        }
    "#,
    &["0123"]
);

kotlin_run_test!(
    test_do_while_mixed_types_in_condition,
    r#"
        fun main() {
            var i = 0
            var sum = 0L
            do {
                sum += i.toLong()
                i += 1
            } while (i.toLong() < 4L)
            println(sum)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_do_while_nested_break,
    r#"
        fun main() {
            var i = 0
            do {
                i += 1
                if (i == 2) break
            } while (i < 4)
            println(i)
        }
    "#,
    &["2"]
);

kotlin_run_test!(
    test_do_while_large_jump,
    r#"
        fun main() {
            var i = 0
            var out = 0
            do {
                out += i
                i = 100
            } while (i < 10)
            println(out)
        }
    "#,
    &["0"]
);

kotlin_run_test!(
    test_do_while_on_function_result,
    r#"
        fun tick(x: Int): Boolean = x < 2
        fun main() {
            var i = 0
            var out = 0
            do {
                out += i
                i++
            } while (tick(i))
            println(out)
            println(i)
        }
    "#,
    &["1", "2"]
);

kotlin_run_test!(
        test_do_while_zero_to_one,
    r#"
        fun main() {
            var i = 0
            val out = kotlin.run {
                i = 1
                i
            }
            println(i)
            println(out)
        }
    "#,
    &["1", "1"]
);

kotlin_run_test!(
    test_do_while_accumulator_with_return,
    r#"
        fun main() {
            var i = 0
            var out = 0
            do {
                if (i > 2) break
                out += i
                i += 1
            } while (true)
            println(out)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_do_while_with_zero_iterations_guard,
    r#"
        fun main() {
            var i = 0
            var out = 0
            do {
                out += 1
                i += 1
            } while (false)
            println(out)
            println(i)
        }
    "#,
    &["1", "1"]
);

kotlin_run_test!(
    test_do_while_with_math_progression,
    r#"
        fun main() {
            var i = 0
            var out = 0
            do {
                out += i * 2
                i += 1
            } while (i < 4)
            println(out)
        }
    "#,
    &["12"]
);

kotlin_run_test!(
    test_do_while_boolean_guard_mix,
    r#"
        fun shouldRun(v: Int): Boolean = v % 2 == 0
        fun main() {
            var i = 0
            var out = ""
            do {
                if (shouldRun(i)) out += i.toString()
                i += 1
            } while (i < 5)
            println(out)
        }
    "#,
    &["02"]
);

kotlin_run_test!(
    test_do_while_with_mutable_flag,
    r#"
        fun main() {
            var i = 0
            var out = ""
            var running = true
            do {
                if (i >= 3) running = false
                out += i.toString()
                i += 1
            } while (running)
            println(out)
        }
    "#,
    &["012"]
);

kotlin_run_test!(
    test_do_while_nested_variable_scope,
    r#"
        fun main() {
            var out = 0
            do {
                var inner = 0
                inner += 1
                out += inner
            } while (out < 3)
            println(out)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_do_while_counter_with_return,
    r#"
        fun main() {
            var i = 0
            var out = 0
            do {
                if (i == 3) {
                    out += 1
                    i = 10
                } else {
                    out += 2
                }
                i += 1
            } while (i < 5)
            println(out)
        }
    "#,
    &["7"]
);

kotlin_run_test!(
    test_do_while_collect_negative,
    r#"
        fun main() {
            var i = -1
            var out = 0
            do {
                out += i
                i -= 1
            } while (i > -4)
            println(out)
        }
    "#,
    &["-6"]
);

kotlin_run_test!(
    test_do_while_string_counter,
    r##"
        fun main() {
            var i = 0
            var out = ""
            do {
                out += "#" + i.toString()
                i += 1
            } while (i < 3)
            println(out)
        }
    "##,
    &["#0#1#2"]
);

kotlin_run_test!(
    test_do_while_in_function,
    r#"
        fun sum(n: Int): Int {
            var i = 0
            var out = 0
            do {
                out += i
                i++
            } while (i < n)
            return out
        }
        fun main() {
            println(sum(4))
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_do_while_long_progression,
    r#"
        fun main() {
            var i = 1L
            var out = 0L
            do {
                out += i
                i += 2
            } while (i < 8)
            println(out)
        }
    "#,
    &["16"]
);
