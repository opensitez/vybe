kotlin_run_test!(
    test_until_empty_range_has_no_elements,
    r#"
        fun main() {
            var sum = 0
            for (i in 4 until 4) sum += i
            println(sum)
        }
    "#,
    &["0"]
);

kotlin_run_test!(
    test_down_to_with_reverse_step,
    r#"
        fun main() {
            val values = 10 downTo 1 step 4
            println(values.toList().joinToString(","))
        }
    "#,
    &["10,6,2"]
);

kotlin_run_test!(
    test_range_contains_boundaries,
    r#"
        fun main() {
            val r = 1..4
            println(1 in r)
            println(4 in r)
            println(5 in r)
        }
    "#,
    &["true", "true", "false"]
);

kotlin_run_test!(
    test_open_ended_range_to_with_negative_start,
    r#"
        fun main() {
            val r = -2..2
            println(r.first)
            println(r.last)
            println(r.count())
        }
    "#,
    &["-2", "2", "5"]
);

kotlin_run_test!(
    test_range_step_one_honors_step_overload,
    r#"
        fun main() {
            val values = (0..10).step(3)
            println(values.toList().joinToString(","))
        }
    "#,
    &["0,3,6,9"]
);

kotlin_run_test!(
    test_char_range_iteration,
    r#"
        fun main() {
            val text = ('a'..'d').joinToString("")
            println(text)
        }
    "#,
    &["abcd"]
);

kotlin_run_test!(
    test_long_range_step_zero_edge,
    r#"
        fun main() {
            val out = (1L..3L step 2).toList()
            println(out.joinToString(","))
        }
    "#,
    &["1,3"]
);

kotlin_run_test!(
    test_range_projection_is_monotonic,
    r#"
        fun main() {
            val r = (0..9)
            val down = r.reversed()
            println(down.first())
            println(down.last())
        }
    "#,
    &["9", "0"]
);

kotlin_run_test!(
    test_empty_down_to_range_yields_none,
    r#"
        fun main() {
            var total = 0
            for (i in 3 downTo 7) {
                total += i
            }
            println(total)
        }
    "#,
    &["0"]
);

kotlin_run_test!(
    test_until_with_negative_step_invariant,
    r#"
        fun main() {
            val values = 5 until 0
            println(values.count())
            println(values.toList().isEmpty())
        }
    "#,
    &["0", "true"]
);

kotlin_run_test!(
    test_range_within_range_contains,
    r#"
        fun main() {
            val outer = 1..20
            val inner = 5..8
            println(inner.first() in outer)
            println(inner.last() in outer)
            println(21 in outer)
        }
    "#,
    &["true", "true", "false"]
);

kotlin_run_test!(
    test_integer_range_to_list_size,
    r#"
        fun main() {
            println((100..102).toList().size)
            println((100 downTo 100).toList().size)
        }
    "#,
    &["3", "1"]
);
