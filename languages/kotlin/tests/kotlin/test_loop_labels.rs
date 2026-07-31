kotlin_run_test!(
    test_outer_break_label_skips_outer_tail,
    r#"
        fun main() {
            var out = 0
            outer@ for (i in 1..4) {
                for (j in 1..4) {
                    if (i == 2) break@outer
                    out += i + j
                }
            }
            println(out)
        }
    "#,
    &["7"]
);

kotlin_run_test!(
    test_outer_continue_label_skips_iteration,
    r#"
        fun main() {
            var out = 0
            outer@ for (i in 1..3) {
                for (j in 1..2) {
                    if (j == 2) continue@outer
                    out += i
                }
            }
            println(out)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_inner_continue_label,
    r#"
        fun main() {
            var out = 0
            for (i in 1..3) {
                inner@ for (j in 1..3) {
                    if (j == 2) continue@inner
                    out += j
                }
            }
            println(out)
        }
    "#,
    &["12"]
);

kotlin_run_test!(
    test_label_on_while_block,
    r#"
        fun main() {
            var i = 0
            var out = 0
            loop@ while (true) {
                i += 1
                if (i == 5) break@loop
                out += i
            }
            println(out)
        }
    "#,
    &["10"]
);

kotlin_run_test!(
    test_do_while_with_label,
    r#"
        fun main() {
            var i = 0
            var out = 0
            mark@ do {
                out += i
                i += 1
            } while (i < 3)
            println(out)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_nested_labeled_while,
    r#"
        fun main() {
            var i = 0
            var out = 0
            outer@ while (i < 3) {
                var j = 0
                while (j < 3) {
                    if (i == 1 && j == 1) break@outer
                    out += i + j
                    j += 1
                }
                i += 1
            }
            println(out)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_continue_label_with_for,
    r#"
        fun main() {
            var out = 0
            outer@ for (i in 1..4) {
                for (j in 1..4) {
                    if (j == 1) continue@outer
                    out += i
                }
            }
            println(out)
        }
    "#,
    &["12"]
);

kotlin_run_test!(
    test_label_for_outer_then_inner,
    r#"
        fun main() {
            var out = ""
            outer@ for (i in 1..3) {
                inner@ for (j in 1..3) {
                    if (i == 2) continue@outer
                    out += i.toString() + j.toString()
                }
            }
            println(out)
        }
    "#,
    &["1112132333"]
);

kotlin_run_test!(
    test_label_before_function_call,
    r#"
        fun tick(): Int = 1
        fun main() {
            var out = 0
            loop@ while (true) {
                if (tick() == 0) continue
                out += 1
                if (out == 4) break@loop
            }
            println(out)
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_label_nested_label_names,
    r#"
        fun main() {
            var out = 0
            one@ for (i in 1..3) {
                two@ for (j in 1..3) {
                    three@ for (k in 1..3) {
                        if (i == 2 && j == 2 && k == 2) break@two
                        out += 1
                    }
                }
            }
            println(out)
        }
    "#,
    &["23"]
);

kotlin_run_test!(
    test_label_continue_in_labeled_for,
    r#"
        fun main() {
            var out = 0
            outer@ for (i in 1..4) {
                for (j in 1..4) {
                    if (j == 3) continue@outer
                    out += j
                }
            }
            println(out)
        }
    "#,
    &["18"]
);

kotlin_run_test!(
    test_label_break_in_repeat,
    r#"
        fun main() {
            var i = 0
            var out = 0
            repeat@ repeat(5) {
                i += 1
                if (i == 3) {
                    out = 99
                    return@repeat
                }
            }
            println(out)
        }
    "#,
    &["99"]
);

kotlin_run_test!(
    test_labelled_do_while_like_while,
    r#"
        fun main() {
            var i = 0
            var out = 0
            block@ while (i < 5) {
                out += i
                i += 1
                continue@block
            }
            println(out)
        }
    "#,
    &["10"]
);

kotlin_run_test!(
    test_label_without_jump,
    r#"
        fun main() {
            var out = 0
            outer@ for (i in 1..2) {
                out += i
            }
            println(out)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_for_with_guarded_labels,
    r#"
        fun isAllowed(x: Int): Boolean = x % 2 == 0
        fun main() {
            var out = 0
            outer@ for (i in 1..6) {
                if (!isAllowed(i)) continue@outer
                out += i
            }
            println(out)
        }
    "#,
    &["12"]
);

kotlin_run_test!(
    test_label_chain_in_nested_for,
    r#"
        fun main() {
            var out = 0
            alpha@ for (i in 1..3) {
                beta@ for (j in 1..3) {
                    if (i + j == 6) break@alpha
                    out += 1
                }
            }
            println(out)
        }
    "#,
    &["7"]
);

kotlin_run_test!(
    test_label_on_while_with_continue,
    r#"
        fun main() {
            var i = 0
            var out = 0
            outer@ while (i < 6) {
                i += 1
                if (i % 2 == 0) continue@outer
                out += i
            }
            println(out)
        }
    "#,
    &["9"]
);

kotlin_run_test!(
    test_label_in_for_each,
    r#"
        fun main() {
            var out = ""
            outer@ for (x in intArrayOf(1,2,3)) {
                for (y in intArrayOf(4,5)) {
                    if (y == 5) continue@outer
                    out += y.toString()
                }
            }
            println(out)
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_label_block_return_compat,
    r#"
        fun main() {
            var out = 0
            one@ for (i in 1..2) {
                two@ for (j in 1..2) {
                    if (i == 2 && j == 2) break@one
                    out += j
                }
            }
            println(out)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_labeled_break_then_continue,
    r#"
        fun main() {
            var out = 0
            for (i in 1..4) {
                block@ for (j in 1..4) {
                    if (j == 2) break@block
                    if (i == 3) continue
                    out += 1
                }
            }
            println(out)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_label_named_loop_in_function,
    r#"
        fun score(x: Int): Int {
            var out = 0
            outer@ for (i in 1..x) {
                if (i == 4) break@outer
                out += i
            }
            return out
        }
        fun main() {
            println(score(6))
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_repeated_labeled_for,
    r#"
        fun main() {
            var out = 0
            for (i in 1..5) {
                mark@ for (j in 1..5) {
                    if (i + j > 6) break@mark
                    out += 1
                }
            }
            println(out)
        }
    "#,
    &["21"]
);

kotlin_run_test!(
    test_labelled_continue_when_none,
    r#"
        fun main() {
            var out = 0
            outer@ for (i in 1..4) {
                if (i == 3) continue@outer
                out += i
            }
            println(out)
        }
    "#,
    &["7"]
);

kotlin_run_test!(
    test_labelled_break_if_nested,
    r#"
        fun main() {
            var out = 0
            outer@ for (i in 1..4) {
                for (j in 1..4) {
                    if (i == 3 && j == 2) break@outer
                    out += 1
                }
            }
            println(out)
        }
    "#,
    &["13"]
);

kotlin_run_test!(
    test_label_expression_result,
    r#"
        fun main() {
            val value = run {
                outer@ for (i in 1..3) {
                    if (i == 2) continue@outer
                }
                42
            }
            println(value)
        }
    "#,
    &["42"]
);

kotlin_run_test!(
    test_nested_label_skip_single_iteration,
    r#"
        fun main() {
            var out = 0
            outer@ for (i in 1..4) {
                for (j in 1..4) {
                    if (j == 3) continue@outer
                    out += j
                }
            }
            println(out)
        }
    "#,
    &["12"]
);

kotlin_run_test!(
    test_label_with_while_and_if,
    r#"
        fun main() {
            var i = 0
            var out = 0
            outer@ while (i < 6) {
                i += 1
                if (i == 4) continue@outer
                if (i == 5) break@outer
                out += i
            }
            println(out)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_double_label_chain,
    r#"
        fun main() {
            var out = 0
            one@ for (i in 1..3) {
                two@ for (j in 1..3) {
                    if (i == 2 && j == 2) continue@one
                    out += i + j
                }
            }
            println(out)
        }
    "#,
    &["10"]
);

kotlin_run_test!(
    test_labelled_continue_in_nested,
    r#"
        fun main() {
            var out = 0
            outer@ for (i in 1..3) {
                for (j in 1..3) {
                    if (i == 2) continue@outer
                    out += j
                }
            }
            println(out)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_label_for_do_while_style,
    r#"
        fun main() {
            var i = 0
            var out = ""
            mark@ for (ch in 1..3) {
                out += ch.toString()
                if (ch == 2) continue@mark
                out += "x"
            }
            println(out)
        }
    "#,
    &["1x2x3x"]
);
