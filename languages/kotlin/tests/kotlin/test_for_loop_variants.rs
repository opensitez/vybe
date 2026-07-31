kotlin_run_test!(
    test_for_over_inclusive_range_sum,
    r#"
        fun main() {
            var total = 0
            for (i in 1..4) {
                total += i
            }
            println(total)
        }
    "#,
    &["10"]
);

kotlin_run_test!(
    test_for_over_until_sum,
    r#"
        fun main() {
            var total = 0
            for (i in 1 until 4) {
                total += i
            }
            println(total)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_for_down_to_range,
    r#"
        fun main() {
            var out = ""
            for (i in 5 downTo 2) {
                out += i.toString()
            }
            println(out)
        }
    "#,
    &["5432"]
);

kotlin_run_test!(
    test_for_range_with_step,
    r#"
        fun main() {
            var total = 0
            for (i in 1..10 step 3) {
                total += i
            }
            println(total)
        }
    "#,
    &["22"]
);

kotlin_run_test!(
    test_for_nested_range_sum,
    r#"
        fun main() {
            var total = 0
            for (i in 1..3) {
                for (j in 1..2) {
                    total += i * j
                }
            }
            println(total)
        }
    "#,
    &["18"]
);

kotlin_run_test!(
    test_for_with_break,
    r#"
        fun main() {
            var total = 0
            for (i in 1..10) {
                if (i == 4) break
                total += i
            }
            println(total)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_for_with_continue,
    r#"
        fun main() {
            var total = 0
            for (i in 1..6) {
                if (i % 2 == 0) continue
                total += i
            }
            println(total)
        }
    "#,
    &["9"]
);

kotlin_run_test!(
    test_for_char_range,
    r#"
        fun main() {
            var out = ""
            for (ch in 'a'..'e') {
                out += ch.toString()
            }
            println(out)
        }
    "#,
    &["abcde"]
);

kotlin_run_test!(
    test_for_array_elements,
    r#"
        fun main() {
            val values = intArrayOf(1, 2, 3)
            var total = 0
            for (v in values) {
                total += v
            }
            println(total)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_for_array_with_index,
    r#"
        fun main() {
            val values = intArrayOf(4, 5, 6)
            var out = 0
            for (i in values.indices) {
                out += values[i]
            }
            println(out)
        }
    "#,
    &["15"]
);

kotlin_run_test!(
    test_for_range_var_bounds,
    r#"
        fun main() {
            var start = 2
            var end = 5
            var total = 0
            for (i in start..end) {
                total += i
            }
            println(total)
        }
    "#,
    &["14"]
);

kotlin_run_test!(
    test_for_range_negative_bounds,
    r#"
        fun main() {
            var total = 0
            for (i in -2..2) {
                total += i
            }
            println(total)
        }
    "#,
    &["0"]
);

kotlin_run_test!(
    test_for_empty_range,
    r#"
        fun main() {
            var total = 0
            for (i in 5 until 5) {
                total += i
            }
            println(total)
        }
    "#,
    &["0"]
);

kotlin_run_test!(
    test_for_singleton_range,
    r#"
        fun main() {
            var total = 0
            for (i in 7..7) {
                total += i
            }
            println(total)
        }
    "#,
    &["7"]
);

kotlin_run_test!(
    test_for_nested_outer_uses,
    r#"
        fun main() {
            var out = ""
            for (row in 1..2) {
                for (col in 1..2) {
                    out += "${row}${col} "
                }
            }
            println(out.trim())
        }
    "#,
    &["11 12 21 22"]
);

kotlin_run_test!(
    test_for_range_long_steps,
    r#"
        fun main() {
            var count = 0L
            for (i in 1L..7L step 2) {
                count += i
            }
            println(count)
        }
    "#,
    &["16"]
);

kotlin_run_test!(
    test_for_range_membership_inside_loop,
    r#"
        fun main() {
            var out = 0
            for (i in 1..10) {
                if (i in 4..6) out += i
            }
            println(out)
        }
    "#,
    &["15"]
);

kotlin_run_test!(
    test_for_shadowing_boundaries,
    r#"
        fun main() {
            var out = 0
            for (i in 1..3) {
                val i = i + 10
                out += i
            }
            println(out)
        }
    "#,
    &["33"]
);

kotlin_run_test!(
    test_for_mutation_and_visibility,
    r#"
        fun main() {
            var values = intArrayOf(1, 2, 3)
            for (i in values.indices) {
                values[i] = values[i] * 2
            }
            println(values[0] + values[1] + values[2])
        }
    "#,
    &["12"]
);

kotlin_run_test!(
    test_for_do_not_run_when_start_gt_end,
    r#"
        fun main() {
            var out = 0
            for (i in 5 downTo 8) {
                out += i
            }
            println(out)
        }
    "#,
    &["0"]
);

kotlin_run_test!(
    test_for_invoke_in_conditional_guard,
    r#"
        fun isOdd(x: Int): Boolean = x % 2 == 1
        fun main() {
            var out = 0
            for (i in 1..10) {
                if (isOdd(i)) out += i
            }
            println(out)
        }
    "#,
    &["25"]
);

kotlin_run_test!(
    test_for_for_expression_output,
    r#"
        fun main() {
            val rows = intArrayOf(1, 2)
            val cols = intArrayOf(10, 20)
            var out = 0
            for (r in rows) {
                for (c in cols) {
                    out += r + c
                }
            }
            println(out)
        }
    "#,
    &["66"]
);

kotlin_run_test!(
    test_for_range_bound_reuse,
    r#"
        fun main() {
            var start = 2
            var end = 6
            var out = 0
            for (i in start until end) {
                out += i
            }
            println(out)
        }
    "#,
    &["14"]
);

kotlin_run_test!(
    test_for_each_char_in_array,
    r#"
        fun main() {
            val values = charArrayOf('a', 'b', 'c')
            var out = ""
            for (c in values) {
                out += c
            }
            println(out)
        }
    "#,
    &["abc"]
);

kotlin_run_test!(
    test_for_char_predicate_filter,
    r#"
        fun main() {
            var out = ""
            for (ch in 'a'..'f') {
                if (ch != 'd') out += ch
            }
            println(out)
        }
    "#,
    &["abcef"]
);

kotlin_run_test!(
    test_for_large_step,
    r#"
        fun main() {
            var total = 0
            for (i in 0..20 step 5) {
                total += i
            }
            println(total)
        }
    "#,
    &["30"]
);

kotlin_run_test!(
    test_for_down_to_with_conditional,
    r#"
        fun main() {
            var out = 0
            for (i in 10 downTo 1) {
                if (i % 4 == 0) out += i
            }
            println(out)
        }
    "#,
    &["14"]
);

kotlin_run_test!(
    test_for_outer_inner_product,
    r#"
        fun main() {
            var out = 0
            for (i in 1..3) {
                for (j in 1..4) {
                    out += if ((i + j) % 2 == 0) 1 else 0
                }
            }
            println(out)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_for_over_while_style_pattern,
    r#"
        fun main() {
            var i = 1
            var out = 0
            for (x in 1..10) {
                if (i > 5) break
                out += x
                i += 1
            }
            println(out)
        }
    "#,
    &["15"]
);

kotlin_run_test!(
    test_for_indexed_array_accumulator,
    r#"
        fun main() {
            val values = intArrayOf(2, 4, 6, 8)
            var out = 0
            for (i in values.indices) {
                if (i % 2 == 0) out += values[i]
            }
            println(out)
        }
    "#,
    &["8"]
);
