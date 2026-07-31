kotlin_run_test!(
    test_break_with_loop_label_stops_outer_loop,
    r#"
        fun main() {
            var total = 0
            outer@ for (i in 0..4) {
                if (i == 4) {
                    break@outer
                }
                if (i == 2) {
                    continue
                }
                total += i
            }
            println(total)
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_continue_with_label_skips_outer_iteration,
    r#"
        fun main() {
            var events = 0
            outer@ for (i in 0..2) {
                for (j in 0..2) {
                    if (i == 1) {
                        continue@outer
                    }
                    events += 1
                }
            }
            println(events)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_break_inner_loop_does_not_leave_outer_loop,
    r#"
        fun main() {
            var acc = 0
            outer@ for (i in 0..2) {
                for (j in 0..2) {
                    if (j == 1) {
                        break
                    }
                    acc += 1
                }
                acc += 10
            }
            println(acc)
        }
    "#,
    &["36"]
);

kotlin_run_test!(
    test_nested_continue_labeling_controls_flow,
    r#"
        fun main() {
            var values = 0
            outer@ for (i in 0..2) {
                for (j in 0..2) {
                    if (i + j < 2) {
                        continue@outer
                    }
                    values += 1
                }
            }
            println(values)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_while_loop_label_break,
    r#"
        fun main() {
            var i = 0
            var total = 0
            repeat@ while (i < 6) {
                i += 1
                if (i % 2 == 0) {
                    continue@repeat
                }
                if (i > 4) {
                    break@repeat
                }
                total += i
            }
            println(total)
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_do_while_label_control,
    r#"
        fun main() {
            var n = 0
            var acc = 0
            run@ do {
                n += 1
                if (n == 3) {
                    continue@run
                }
                acc += n
            } while (n < 5)
            println(acc)
        }
    "#,
    &["12"]
);

kotlin_run_test!(
    test_multiple_labels_targeted_explicitly,
    r#"
        fun main() {
            var x = 0
            outer@ for (i in 0..1) {
                inner@ for (j in 0..2) {
                    if (i == 1 && j == 0) {
                        continue@outer
                    }
                    if (j == 2) {
                        break@inner
                    }
                    x += 1
                }
            }
            println(x)
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_label_skips_only_after_conditions,
    r#"
        fun main() {
            var out = ""
            outer@ for (ch in listOf('a', 'b', 'c')) {
                for (digit in listOf('1', '2', '3')) {
                    if (ch == 'b' && digit == '2') {
                        continue@outer
                    }
                    out += "$ch$digit|"
                }
            }
            println(out)
        }
    "#,
    &["a1|a2|a3|c1|c2|c3|"]
);

kotlin_run_test!(
    test_label_on_when_like_control_is_invalid,
    r#"
        fun main() {
            var count = 0
            search@ for (n in listOf(1, 2, 3, 4)) {
                val labelValue = if (n == 3) {
                    continue@search
                } else if (n == 4) {
                    break@search
                } else {
                    n
                }
                count += labelValue
            }
            println(count)
        }
    "#,
    &["3"]
);
