use crate::helpers::run_prints;

#[test]
fn test_for_range_inclusive_counts_all_values() {
    let out = run_prints(
        r#"
        fun main() {
            var total = 0
            for (i in 1..4) {
                total += i
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_for_range_until_excludes_upper_bound() {
    let out = run_prints(
        r#"
        fun main() {
            var total = 0
            for (i in 1 until 4) {
                total += i
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_for_range_descending_iterates_in_reverse() {
    let out = run_prints(
        r#"
        fun main() {
            var output = ""
            for (i in 5 downTo 2) {
                output += i.toString()
            }
            println(output)
        }
    "#,
    );
    assert_eq!(out, &["5432"]);
}

#[test]
fn test_for_range_with_step_skips_values() {
    let out = run_prints(
        r#"
        fun main() {
            var output = 0
            for (i in 1..10 step 3) {
                output += i
            }
            println(output)
        }
    "#,
    );
    assert_eq!(out, &["22"]);
}

#[test]
fn test_for_range_descending_with_step_skips_values() {
    let out = run_prints(
        r#"
        fun main() {
            var output = 0
            for (i in 10 downTo 1 step 3) {
                output += i
            }
            println(output)
        }
    "#,
    );
    assert_eq!(out, &["18"]);
}

#[test]
fn test_for_range_empty_when_until_has_no_room() {
    let out = run_prints(
        r#"
        fun main() {
            var sum = 0
            for (i in 3 until 3) {
                sum += i
            }
            println(sum)
        }
    "#,
    );
    assert_eq!(out, &["0"]);
}

#[test]
fn test_for_range_singleton_is_still_executed() {
    let out = run_prints(
        r#"
        fun main() {
            var count = 0
            for (i in 7 downTo 7) {
                count += i
            }
            println(count)
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_for_loop_skips_empty_down_to_range() {
    let out = run_prints(
        r#"
        fun main() {
            var seen = 0
            for (i in 1 downTo 3) {
                seen += i
            }
            println(seen)
        }
    "#,
    );
    assert_eq!(out, &["0"]);
}

#[test]
fn test_for_loop_over_char_range() {
    let out = run_prints(
        r#"
        fun main() {
            var text = ""
            for (c in 'a'..'d') {
                text += c
            }
            println(text)
        }
    "#,
    );
    assert_eq!(out, &["abcd"]);
}

#[test]
fn test_while_loop_runs_until_false() {
    let out = run_prints(
        r#"
        fun main() {
            var i = 1
            var out = 0
            while (i <= 4) {
                out += i
                i += 1
            }
            println(out)
        }
    "#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_while_loop_with_break_stops_immediately() {
    let out = run_prints(
        r#"
        fun main() {
            var i = 0
            var out = 0
            while (true) {
                if (i == 3) break
                out += i
                i += 1
            }
            println(out)
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_while_loop_continue_skips_selected_iteration() {
    let out = run_prints(
        r#"
        fun main() {
            var i = 0
            var total = 0
            while (i < 6) {
                i += 1
                if (i % 2 == 0) continue
                total += i
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_do_while_executes_once_when_condition_false() {
    let out = run_prints(
        r#"
        fun main() {
            var total = 0
            var i = 0
            do {
                total += i
                i += 1
            } while (false)
            println(total)
            println(i)
        }
    "#,
    );
    assert_eq!(out, &["0", "1"]);
}

#[test]
fn test_do_while_executes_until_condition_fails() {
    let out = run_prints(
        r#"
        fun main() {
            var total = 0
            var i = 1
            do {
                total += i
                i += 1
            } while (i <= 4)
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_nested_for_break_affects_inner_loop_only() {
    let out = run_prints(
        r#"
        fun main() {
            var total = 0
            for (i in 1..3) {
                for (j in 1..5) {
                    if (j == 4) break
                    total += i * j
                }
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["36"]);
}

#[test]
fn test_nested_for_with_continue_affects_inner_loop() {
    let out = run_prints(
        r#"
        fun main() {
            var total = 0
            for (i in 1..3) {
                for (j in 1..5) {
                    if (j == 3) continue
                    total += i + j
                }
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["60"]);
}

#[test]
fn test_labeled_break_exits_outer_loop() {
    let out = run_prints(
        r#"
        fun main() {
            var total = 0
            outer@ for (i in 1..5) {
                for (j in 1..5) {
                    total += i + j
                    if (i == 3 && j == 2) break@outer
                }
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["21"]);
}

#[test]
fn test_labeled_continue_skips_outer_iteration() {
    let out = run_prints(
        r#"
        fun main() {
            var total = 0
            outer@ for (i in 1..4) {
                for (j in 1..3) {
                    if (j == 3 && i == 2) continue@outer
                    total += i * j
                }
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["42"]);
}

#[test]
fn test_while_with_labeled_outer_continue_and_break() {
    let out = run_prints(
        r#"
        fun main() {
            var i = 0
            var sum = 0
            outer@ while (i < 10) {
                i += 1
                if (i == 3) continue@outer
                if (i == 8) break@outer
                sum += i
            }
            println(sum)
        }
    "#,
    );
    assert_eq!(out, &["27"]);
}

#[test]
fn test_for_over_array_indexes() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(10, 20, 30)
            var sum = 0
            for (i in nums.indices) {
                sum += i + nums[i]
            }
            println(sum)
        }
    "#,
    );
    assert_eq!(out, &["33"]);
}

#[test]
fn test_for_over_array_elements_uses_current_values() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(2, 4, 6)
            var total = 0
            for (n in nums) {
                total += n
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["12"]);
}

#[test]
fn test_for_over_list_of_strings() {
    let out = run_prints(
        r#"
        fun main() {
            val words = listOf("a", "b", "c")
            println(words.joinToString(","))
            var joined = ""
            for (word in words) {
                joined += word
            }
            println(joined)
        }
    "#,
    );
    assert_eq!(out, &["a,b,c", "abc"]);
}

#[test]
fn test_for_over_string_iterates_code_points() {
    let out = run_prints(
        r#"
        fun main() {
            var out = ""
            for (ch in "K1") {
                out += ch.uppercase()
            }
            println(out)
        }
    "#,
    );
    assert_eq!(out, &["K1"]);
}

#[test]
fn test_for_over_map_entries_collects_keys_and_values() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mapOf("a" to 1, "b" to 2)
            var keys = ""
            var sum = 0
            for (entry in values.entries) {
                keys += entry.key
                sum += entry.value
            }
            println(keys)
            println(sum)
        }
    "#,
    );
    assert_eq!(out, &["ab", "3"]);
}

#[test]
fn test_for_with_destructuring_map_pairs() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mapOf("x" to 3, "y" to 4)
            var total = 0
            for ((key, value) in values) {
                if (key == "x") {
                    total += value
                } else {
                    total += value * 2
                }
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["11"]);
}

#[test]
fn test_for_in_range_inside_while_interaction() {
    let out = run_prints(
        r#"
        fun main() {
            var rounds = 0
            while (rounds < 2) {
                rounds += 1
                var row = 0
                for (i in 1..3) {
                    row += i
                }
                println(row)
            }
        }
    "#,
    );
    assert_eq!(out, &["6", "6"]);
}

#[test]
fn test_repeat_loops_like_control_structure() {
    let out = run_prints(
        r#"
        fun main() {
            var total = 0
            repeat(4) {
                total += 1
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["4"]);
}

#[test]
fn test_control_flow_with_nested_do_while_and_for() {
    let out = run_prints(
        r#"
        fun main() {
            var rounds = 0
            var total = 0
            do {
                for (i in 1..3) {
                    total += i
                }
                rounds += 1
            } while (rounds < 2)
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["12"]);
}

#[test]
fn test_for_on_reversed_range_uses_expected_order_for_nested_accumulation() {
    let out = run_prints(
        r#"
        fun main() {
            var sequence = ""
            for (i in 3 downTo 1) {
                for (j in 0..1) {
                    sequence += i
                    sequence += ":"
                }
            }
            println(sequence)
        }
    "#,
    );
    assert_eq!(out, &["3:3:2:2:1:1:"]);
}

#[test]
fn test_while_condition_recomputed_each_iteration() {
    let out = run_prints(
        r#"
        fun main() {
            var i = 0
            var threshold = 3
            var total = 0
            while (i < threshold) {
                total += 1
                threshold += if (i == 1) 2 else 0
                i += 1
            }
            println(i)
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["4", "4"]);
}

#[test]
fn test_break_and_continue_with_nested_while_loops() {
    let out = run_prints(
        r#"
        fun main() {
            var i = 0
            var outerTotal = 0
            while (i < 3) {
                var j = 0
                while (j < 4) {
                    j += 1
                    if (j == 2) continue
                    if (i == 1 && j == 4) break
                    outerTotal += i + j
                }
                i += 1
            }
            println(outerTotal)
        }
    "#,
    );
    assert_eq!(out, &["15"]);
}

#[test]
fn test_while_loop_with_boolean_flag_and_loop_variable() {
    let out = run_prints(
        r#"
        fun main() {
            var active = true
            var i = 0
            var total = 0
            while (active) {
                total += i
                i += 1
                active = i < 5
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["10"]);
}

#[test]
fn test_for_over_empty_collection_keeps_accumulator_zero() {
    let out = run_prints(
        r#"
        fun main() {
            val values = intArrayOf()
            var seen = 0
            for (value in values) {
                seen += value
            }
            println(seen)
        }
    "#,
    );
    assert_eq!(out, &["0"]);
}

#[test]
fn test_repeat_zero_iteration_is_noop() {
    let out = run_prints(
        r#"
        fun main() {
            var total = 0
            repeat(0) {
                total += 1
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["0"]);
}

#[test]
fn test_while_labeled_continue_targets_outer_loop() {
    let out = run_prints(
        r#"
        fun main() {
            var i = 0
            var total = 0
            outer@ while (i < 3) {
                i += 1
                var j = 0
                while (j < 3) {
                    j += 1
                    if (j == 2) continue@outer
                    total += i * j
                }
                total += 10
            }
            println(i)
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["3", "6"]);
}

#[test]
fn test_do_while_runs_once_before_condition_is_checked() {
    let out = run_prints(
        r#"
        fun main() {
            var total = 0
            var i = 9
            do {
                total += i
                i -= 2
            } while (i < 0)
            println(total)
            println(i)
        }
    "#,
    );
    assert_eq!(out, &["9", "7"]);
}

#[test]
fn test_for_on_char_range_with_step_includes_expected_codepoints() {
    let out = run_prints(
        r#"
        fun main() {
            var text = ""
            for (c in 'a'..'f' step 2) {
                text += c.toString()
            }
            println(text)
        }
    "#,
    );
    assert_eq!(out, &["ace"]);
}
