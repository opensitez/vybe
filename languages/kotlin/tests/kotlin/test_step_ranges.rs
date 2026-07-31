kotlin_run_test!(
    test_down_to_progression_count,
    r#"
        fun main() {
            val values = (5 downTo 1).toList()
            println(values.joinToString(","))
        }
    "#,
    &["5,4,3,2,1"]
);

kotlin_run_test!(
    test_until_is_exclusive_end,
    r#"
        fun main() {
            val values = (1 until 4).toList()
            println(values.joinToString(";"))
        }
    "#,
    &["1;2;3"]
);

kotlin_run_test!(
    test_step_by_skip,
    r#"
        fun main() {
            val values = (1..10 step 3).toList()
            println(values.joinToString(","))
        }
    "#,
    &["1,4,7,10"]
);

kotlin_run_test!(
    test_down_to_step_skips,
    r#"
        fun main() {
            val values = (10 downTo 1 step 4).toList()
            println(values.joinToString(","))
        }
    "#,
    &["10,6,2"]
);

kotlin_run_test!(
    test_range_reversed_after_to_list,
    r#"
        fun main() {
            val values = (1..5).toList().asReversed()
            println(values.joinToString(","))
        }
    "#,
    &["5,4,3,2,1"]
);

kotlin_run_test!(
    test_char_range,
    r#"
        fun main() {
            val values = ('a'..'e').toList()
            println(values.joinToString(""))
            println(('e' downTo 'c').toList().joinToString(""))
        }
    "#,
    &["abcde", "edc"]
);

kotlin_run_test!(
    test_range_contains_inclusive_bounds,
    r#"
        fun main() {
            println(1 in 1..3)
            println(3 in 1..3)
            println(4 in 1..3)
        }
    "#,
    &["true", "true", "false"]
);

kotlin_run_test!(
    test_until_range_exclusive,
    r#"
        fun main() {
            val values = (1 until 1).toList()
            val other = (3 until 5).toList()
            println(values.isEmpty())
            println(other.joinToString(","))
        }
    "#,
    &["true", "3,4"]
);

kotlin_run_test!(
    test_step_with_while_like_loop,
    r#"
        fun main() {
            var count = 0
            for (v in 0..10 step 2) {
                count += v
            }
            println(count)
        }
    "#,
    &["30"]
);

kotlin_run_test!(
    test_down_to_on_negative_numbers,
    r#"
        fun main() {
            println((-1 downTo -3).toList().joinToString(","))
            println((-3..-1).toList().joinToString(","))
        }
    "#,
    &["-1,-2,-3", "-3,-2,-1"]
);

kotlin_run_test!(
    test_range_iteration_edge_cases,
    r#"
        fun main() {
            var total = 0
            for (x in 0 until 1) {
                total += x
            }
            var done = 0
            for (x in 2 downTo 2) {
                done += x
            }
            println(total)
            println(done)
        }
    "#,
    &["0", "2"]
);
