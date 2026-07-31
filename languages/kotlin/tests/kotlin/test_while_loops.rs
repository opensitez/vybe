kotlin_run_test!(
    test_while_loop_basic_sum,
    r#"
        fun main() {
            var i = 0
            var total = 0
            while (i < 5) {
                total += i
                i += 1
            }
            println(total)
        }
    "#,
    &["10"]
);

kotlin_run_test!(
    test_while_loop_with_break,
    r#"
        fun main() {
            var i = 0
            var out = 0
            while (i < 10) {
                i += 1
                if (i == 4) break
                out += i
            }
            println(i)
            println(out)
        }
    "#,
    &["4", "6"]
);

kotlin_run_test!(
    test_while_loop_with_continue,
    r#"
        fun main() {
            var i = 0
            var out = 0
            while (i < 6) {
                i += 1
                if (i % 2 == 0) continue
                out += i
            }
            println(out)
        }
    "#,
    &["9"]
);

kotlin_run_test!(
    test_do_while_executes_once_when_false,
    r#"
        fun main() {
            var i = 0
            val out = do {
                i += 1
                i
            } while (i > 10)
            println(out)
        }
    "#,
    &["1"]
);

kotlin_run_test!(
    test_nested_while_depth,
    r#"
        fun main() {
            var i = 0
            var total = 0
            while (i < 3) {
                var j = 0
                while (j < 2) {
                    total += i + j
                    j += 1
                }
                i += 1
            }
            println(total)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_while_true_with_break_condition,
    r#"
        fun main() {
            var i = 0
            var total = 0
            while (true) {
                if (i == 4) break
                total += i
                i += 1
            }
            println(total)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_do_while_accumulator,
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
    test_while_post_condition,
    r#"
        fun main() {
            var i = 5
            var out = 0
            while (i >= 1) {
                out += i
                i -= 2
            }
            println(out)
        }
    "#,
    &["9"]
);

kotlin_run_test!(
    test_while_with_string_accumulator,
    r#"
        fun main() {
            var i = 0
            var text = ""
            while (i < 3) {
                text += i.toString()
                i += 1
            }
            println(text)
        }
    "#,
    &["012"]
);

kotlin_run_test!(
    test_continue_in_nested_while,
    r#"
        fun main() {
            var i = 0
            var out = 0
            while (i < 5) {
                i += 1
                var inner = 0
                while (inner < i) {
                    inner += 1
                    if (inner == 1) continue
                    out += 1
                    break
                }
            }
            println(out)
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_do_while_with_early_break,
    r#"
        fun main() {
            var i = 0
            var out = 0
            do {
                if (i == 3) break
                out += i
                i += 1
            } while (true)
            println(out)
        }
    "#,
    &["3"]
);
