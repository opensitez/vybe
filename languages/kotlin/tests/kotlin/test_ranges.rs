use crate::helpers::run_prints;

#[test]
fn test_inclusive_range_iteration() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            for (i in 1..4) {
                total += i
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["10"]);
}

#[test]
fn test_until_range_excludes_upper_bound() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            for (i in 1 until 4) {
                total += i
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_down_to_range_iteration() {
    let out = run_prints(r#"
        fun main() {
            var output = ""
            for (i in 5 downTo 3) {
                output += i.toString()
            }
            println(output)
        }
    "#);
    assert_eq!(out, &["543"]);
}

#[test]
fn test_range_with_step() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            for (i in 1..10 step 3) {
                total += i
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["22"]);
}

#[test]
fn test_range_membership_true_false() {
    let out = run_prints(r#"
        fun main() {
            val x = 4
            val y = 8
            println(x in 1..6)
            println(y in 1 until 6)
            println(2 in 6 downTo 1)
        }
    "#);
    assert_eq!(out, &["true", "false", "true"]);
}

#[test]
fn test_nested_ranges() {
    let out = run_prints(r#"
        fun main() {
            var count = 0
            for (row in 1..2) {
                for (col in row..3) {
                    count += col
                }
            }
            println(count)
        }
    "#);
    assert_eq!(out, &["10"]);
}

#[test]
fn test_empty_like_range_semantics() {
    let out = run_prints(r#"
        fun main() {
            var seen = 0
            for (i in 3 until 3) {
                seen += i
            }
            println(seen)
        }
    "#);
    assert_eq!(out, &["0"]);
}

#[test]
fn test_range_derived_bounds_in_function() {
    let out = run_prints(r#"
        fun sumInRange(start: Int, end: Int): Int {
            var total = 0
            for (value in start..end) {
                total += value
            }
            return total
        }

        fun main() {
            println(sumInRange(3, 5))
        }
    "#);
    assert_eq!(out, &["12"]);
}

#[test]
fn test_range_step_with_start_end_expressions() {
    let out = run_prints(r#"
        fun build(start: Int, end: Int, step: Int): String {
            var out = ""
            for (value in start..end step step) {
                out += value.toString()
            }
            return out
        }

        fun main() {
            println(build(1, 7, 2))
        }
    "#);
    assert_eq!(out, &["1357"]);
}

#[test]
fn test_inclusive_singleton_range() {
    let out = run_prints(r#"
        fun main() {
            var output = ""
            for (value in 3..3) {
                output += value.toString()
            }
            println(output)
            println(3 in 3..3)
            println(4 in 3..3)
        }
    "#);
    assert_eq!(out, &["3", "true", "false"]);
}

#[test]
fn test_down_to_singleton_range() {
    let out = run_prints(r#"
        fun main() {
            var output = ""
            for (value in 3 downTo 3) {
                output += value.toString()
            }
            println(output)
            println(3 in 3 downTo 3)
            println(2 in 3 downTo 3)
        }
    "#);
    assert_eq!(out, &["3", "true", "false"]);
}

#[test]
fn test_descending_range_with_empty_bounds() {
    let out = run_prints(r#"
        fun main() {
            var count = 0
            for (value in 2 downTo 5) {
                count += value
            }
            println(count)
            println(5 in 2 downTo 5)
        }
    "#);
    assert_eq!(out, &["0", "false"]);
}

#[test]
fn test_negative_numbers_in_inclusive_range() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            for (value in -2..2) {
                total += value
            }
            println(total)
            println(-1 in -2..2)
            println(3 in -2..2)
        }
    "#);
    assert_eq!(out, &["0", "true", "false"]);
}

#[test]
fn test_char_range_iteration_and_membership() {
    let out = run_prints(r#"
        fun main() {
            var output = ""
            for (value in 'a'..'d') {
                output += value.toString()
            }
            println(output)
            println('b' in 'a'..'d')
            println('x' in 'a'..'d')
        }
    "#);
    assert_eq!(out, &["abcd", "true", "false"]);
}

#[test]
fn test_char_range_with_step() {
    let out = run_prints(r#"
        fun main() {
            var output = ""
            for (value in 'a'..'f' step 2) {
                output += value.toString()
            }
            println(output)
            println('e' in 'a'..'f' step 2)
            println('d' in 'a'..'f' step 2)
        }
    "#);
    assert_eq!(out, &["ace", "true", "false"]);
}

#[test]
fn test_long_range_iteration() {
    let out = run_prints(r#"
        fun main() {
            var total = 0L
            for (value in 1L..7L step 2) {
                total += value
            }
            println(total)
            println(6L in 1L..7L)
            println(8L in 1L until 8L)
        }
    "#);
    assert_eq!(out, &["16", "true", "false"]);
}

#[test]
fn test_range_is_empty_knowledge() {
    let out = run_prints(r#"
        fun main() {
            println((1..0).isEmpty())
            println((0 until 0).isEmpty())
            println((5 downTo 10).isEmpty())
        }
    "#);
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_range_first_and_last_properties() {
    let out = run_prints(r#"
        fun main() {
            val growing = 1..4
            val declining = 4 downTo 1
            println(growing.first)
            println(growing.last)
            println(declining.first)
            println(declining.last)
        }
    "#);
    assert_eq!(out, &["1", "4", "4", "1"]);
}

#[test]
fn test_range_membership_with_offset_bounds() {
    let out = run_prints(r#"
        fun main() {
            val lower = 2
            val upper = 5
            var result = ""
            for (value in (lower + 1)..upper) {
                result += value.toString()
            }
            println(result)
            println((lower + 1) in (lower + 1)..upper)
            println((upper) in (lower + 1)..upper)
        }
    "#);
    assert_eq!(out, &["34", "true", "false"]);
}

#[test]
fn test_range_in_loop_with_break_condition() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            for (value in 1..20) {
                if (value == 7) {
                    break
                }
                total += value
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["21"]);
}

#[test]
fn test_range_in_loop_with_continue_condition() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            for (value in 1..10) {
                if (value % 2 == 0) {
                    continue
                }
                total += value
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["25"]);
}

#[test]
fn test_nested_range_product() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            for (a in 1..3) {
                for (b in a..4) {
                    total += a * b
                }
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["34"]);
}

#[test]
fn test_range_based_map_like_aggregation() {
    let out = run_prints(r#"
        fun build(start: Int, end: Int, step: Int): Int {
            var total = 1
            for (value in start..end step step) {
                total *= value
            }
            return total
        }

        fun main() {
            println(build(1, 4, 1))
            println(build(2, 6, 2))
        }
    "#);
    assert_eq!(out, &["24", "48"]);
}

#[test]
fn test_reversed_range_step_two() {
    let out = run_prints(r#"
        fun main() {
            var result = ""
            for (value in 10 downTo 1 step 3) {
                result += value.toString()
            }
            println(result)
        }
    "#);
    assert_eq!(out, &["10976"]);
}

#[test]
fn test_range_with_zero_like_start() {
    let out = run_prints(r#"
        fun main() {
            var output = ""
            for (value in 0..3) {
                output += value.toString()
            }
            println(output)
            println(0 in 0 until 3)
            println(3 in 0 until 3)
        }
    "#);
    assert_eq!(out, &["0123", "true", "false"]);
}

#[test]
fn test_unbounded_step_expression_range() {
    let out = run_prints(r#"
        fun main() {
            val step = 2
            var total = 0
            for (value in 1..7 step step) {
                total += value
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["16"]);
}

#[test]
fn test_range_in_function_argument() {
    let out = run_prints(r#"
        fun containsTarget(range: IntRange, target: Int): Boolean {
            return target in range
        }

        fun main() {
            println(containsTarget(3..8, 5))
            println(containsTarget(3..8, 9))
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_range_sum_with_offset_function() {
    let out = run_prints(r#"
        fun span(start: Int, length: Int): IntRange {
            return start..(start + length)
        }

        fun main() {
            var total = 0
            for (value in span(2, 3)) {
                total += value
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["14"]);
}

#[test]
fn test_range_equality_and_string_representation() {
    let out = run_prints(r#"
        fun main() {
            val rangeA = 1..3
            val rangeB = 1..3
            val rangeC = 1..4
            println(rangeA == rangeB)
            println(rangeA == rangeC)
            println(rangeA.toString())
        }
    "#);
    assert_eq!(out, &["true", "false", "1..3"]);
}

#[test]
fn test_range_step_property_is_intprogression_value() {
    let out = run_prints(r#"
        fun main() {
            val range = 1..10 step 3
            println(range.step)
            println(range.first)
            println(range.last)
        }
    "#);
    assert_eq!(out, &["3", "1", "10"]);
    // The last for 1..10 step 3 is 10 because progression includes start/end boundary.
}

#[test]
fn test_reversed_range_iteration_order_and_bounds() {
    let out = run_prints(r#"
        fun main() {
            val forward = (1..7).reversed()
            var forwardFirst = ""
            for (value in forward) {
                forwardFirst += value.toString()
            }
            val reversed = (7 downTo 1).reversed()
            var reversedFirst = ""
            for (value in reversed) {
                reversedFirst += value.toString()
            }
            println(forward.first())
            println(forward.last())
            println(reversedFirst)
        }
    "#);
    assert_eq!(out, &["7", "1", "1234567"]);
}

#[test]
fn test_range_size_and_empty_count_contract() {
    let out = run_prints(r#"
        fun main() {
            val full = 1..6
            val down = 6 downTo 1
            val empty = 1..0
            println(full.count())
            println(down.count())
            println(empty.count())
        }
    "#);
    assert_eq!(out, &["6", "6", "0"]);
}

#[test]
fn test_range_to_set_and_to_list_semantics() {
    let out = run_prints(r#"
        fun main() {
            val list = (1..4).toList()
            val set = (1..4).toSet()
            println(list.size)
            println(list.joinToString())
            println(set.size)
            println(set.contains(4))
        }
    "#);
    assert_eq!(out, &["4", "1, 2, 3, 4", "4", "true"]);
}

#[test]
fn test_coerce_in_range_boundaries() {
    let out = run_prints(r#"
        fun main() {
            val allowed = 3..8
            println(1.coerceIn(allowed))
            println(5.coerceIn(allowed))
            println(9.coerceIn(allowed))
        }
    "#);
    assert_eq!(out, &["3", "5", "8"]);
}

#[test]
fn test_char_range_count_and_reverse_view() {
    let out = run_prints(r#"
        fun main() {
            val span = 'b'..'f'
            println(span.count())
            val reversed = span.reversed()
            println(reversed.count())
            println(reversed.first())
            println(reversed.last())
        }
    "#);
    assert_eq!(out, &["5", "5", "f", "b"]);
}

#[test]
fn test_long_range_step_and_contains() {
    let out = run_prints(r#"
        fun main() {
            val timeline = 1000L downTo 995L step 2
            println(timeline.first)
            println(timeline.last)
            println(999L in timeline)
            println(998L in timeline)
            println(timeline.count())
        }
    "#);
    assert_eq!(out, &["1000", "996", "true", "false", "3"]);
}

#[test]
fn test_range_with_step_one_is_equivalent_to_plain_range() {
    let out = run_prints(r#"
        fun main() {
            var a = ""
            for (value in 2..8 step 1) {
                a += value.toString()
            }
            var b = ""
            for (value in 2..8) {
                b += value.toString()
            }
            println(a)
            println(b)
            println(a == b)
        }
    "#);
    assert_eq!(out, &["2345678", "2345678", "true"]);
}

#[test]
fn test_negative_then_positive_step_is_empty() {
    let out = run_prints(r#"
        fun main() {
            val broken = (1..5).step(-2)
            println(broken.isEmpty())
            println(broken.count())
            println(1 in broken)
        }
    "#);
    assert_eq!(out, &["true", "0", "false"]);
}

#[test]
fn test_range_contains_uses_open_ended_bounds_expressions() {
    let out = run_prints(r#"
        fun clamp(value: Int, base: Int, width: Int): Boolean {
            return value in (base until (base + width))
        }

        fun main() {
            println(clamp(3, 0, 3))
            println(clamp(2, -2, 4))
            println(clamp(3, -2, 4))
        }
    "#);
    assert_eq!(out, &["true", "true", "false"]);
}

#[test]
fn test_range_bounds_are_evaluated_once() {
    let out = run_prints(r#"
        var leftBoundCalls = 0
        var rightBoundCalls = 0

        fun left(): Int {
            leftBoundCalls += 1
            return 1
        }

        fun right(): Int {
            rightBoundCalls += 1
            return 4
        }

        fun main() {
            var total = 0
            for (value in left()..right()) {
                total += value
            }
            println(leftBoundCalls)
            println(rightBoundCalls)
            println(total)
        }
    "#);
    assert_eq!(out, &["1", "1", "10"]);
}

#[test]
fn test_reversed_range_iteration_matches_original_elements() {
    let out = run_prints(r#"
        fun main() {
            var ascending = ""
            for (value in (5 downTo 1).reversed()) {
                ascending += value.toString()
            }
            println(ascending)
            println((5 downTo 1).contains(3))
            println((5 downTo 1).contains(0))
        }
    "#);
    assert_eq!(out, &["12345", "true", "false"]);
}

#[test]
fn test_range_progression_with_stride_without_full_division() {
    let out = run_prints(r#"
        fun main() {
            var values = ""
            for (value in 0..11 step 4) {
                values += value.toString()
                if (value > 8) {
                    values += "|"
                }
            }
            println(values)
            println(10 in (0..11 step 4))
            println(8 in (0..11 step 4))
        }
    "#);
    assert_eq!(out, &["048|", "false", "true"]);
}

#[test]
fn test_range_to_list_and_set_have_snapshot_semantics() {
    let out = run_prints(r#"
        fun main() {
            val values = (1..4).toList()
            val set = (1..4).toMutableList()
            set[0] = 9
            println(values[0])
            println(set[0])
            println((1..4).toMutableList()[0])
        }
    "#);
    assert_eq!(out, &["1", "9", "1"]);
}
