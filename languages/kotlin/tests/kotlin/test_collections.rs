use crate::helpers::run_prints;

#[test]
fn test_array_of_values_and_size() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(1, 2, 3, 4)
            println(nums.size)
            println(nums[0])
            println(nums[3])
        }
    "#,
    );
    assert_eq!(out, &["4", "1", "4"]);
}

#[test]
fn test_array_mutation_in_place() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(10, 20, 30)
            nums[1] = 99
            println(nums[1])
            println(nums[0] + nums[1] + nums[2])
        }
    "#,
    );
    assert_eq!(out, &["99", "139"]);
}

#[test]
fn test_array_iteration_accumulation() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(1, 2, 3)
            var sum = 0
            for (n in nums) {
                sum += n
            }
            println(sum)
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_array_with_indices_loop() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(4, 9, 2)
            var output = 0
            for (i in nums.indices) {
                output += nums[i] * i
            }
            println(output)
        }
    "#,
    );
    assert_eq!(out, &["26"]);
}

#[test]
fn test_two_dimensional_array_access() {
    let out = run_prints(
        r#"
        fun main() {
            val grid = arrayOf(
                arrayOf(1, 2),
                arrayOf(3, 4),
            )
            println(grid[0][1] + grid[1][0])
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_array_of_nullable_slots() {
    let out = run_prints(
        r#"
        fun main() {
            val slots: Array<Int?> = arrayOf(null, null, null)
            slots[1] = 7
            println(slots[0] == null)
            println(slots[1] + 1)
            println(slots[2] == null)
        }
    "#,
    );
    assert_eq!(out, &["true", "8", "true"]);
}

#[test]
fn test_pair_array_destructuring_loop() {
    let out = run_prints(
        r#"
        fun main() {
            val entries = arrayOf(Pair("left", 1), Pair("right", 2))
            var leftTotal = 0
            var rightTotal = 0
            for ((name, value) in entries) {
                if (name == "left") {
                    leftTotal = value
                } else {
                    rightTotal = value
                }
            }
            println(leftTotal)
            println(rightTotal)
        }
    "#,
    );
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_explicit_typed_array() {
    let out = run_prints(
        r#"
        fun main() {
            val nums: Array<Int> = arrayOf(5, 6, 7)
            println(nums.size)
            println(nums[2])
        }
    "#,
    );
    assert_eq!(out, &["3", "7"]);
}

#[test]
fn test_empty_array_behavior() {
    let out = run_prints(
        r#"
        fun main() {
            val empty = arrayOf<Int>()
            println(empty.size)
            if (empty.size == 0) {
                println("empty")
            }
        }
    "#,
    );
    assert_eq!(out, &["0", "empty"]);
}

#[test]
fn test_array_of_factory_by_index() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = Array(4) { idx -> idx + 1 }
            println(nums[0])
            println(nums[3])
            println(nums.size)
        }
    "#,
    );
    assert_eq!(out, &["1", "4", "4"]);
}

#[test]
fn test_array_clone_reference_semantics() {
    let out = run_prints(
        r#"
        fun main() {
            val original = arrayOf(1, 2, 3)
            val shared = original
            shared[1] = 99
            println(original[1])
            println(shared[1])
        }
    "#,
    );
    assert_eq!(out, &["99", "99"]);
}

#[test]
fn test_array_manual_copy_is_independent() {
    let out = run_prints(
        r#"
        fun main() {
            val source = arrayOf(1, 2, 3)
            val copy = Array(source.size) { index -> source[index] }
            copy[0] = 9
            println(source[0])
            println(copy[0])
        }
    "#,
    );
    assert_eq!(out, &["1", "9"]);
}

#[test]
fn test_array_of_nullable_length_three() {
    let out = run_prints(
        r#"
        fun main() {
            val values: Array<Int?> = Array(3) { null }
            values[2] = 14
            println(values[0] == null)
            println(values[2] + 1)
        }
    "#,
    );
    assert_eq!(out, &["true", "15"]);
}

#[test]
fn test_array_swap_in_place() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(1, 2, 3, 4)
            val tmp = nums[0]
            nums[0] = nums[3]
            nums[3] = tmp
            println(nums[0] + nums[3])
            println(nums[1] + nums[2])
        }
    "#,
    );
    assert_eq!(out, &["5", "5"]);
}

#[test]
fn test_two_dimensional_manual_sum() {
    let out = run_prints(
        r#"
        fun main() {
            val grid = arrayOf(
                arrayOf(1, 2, 3),
                arrayOf(4, 5, 6)
            )
            var total = 0
            for (row in grid) {
                for (cell in row) {
                    total += cell
                }
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["21"]);
}

#[test]
fn test_array_diagonal_access() {
    let out = run_prints(
        r#"
        fun main() {
            val grid = arrayOf(
                arrayOf(9, 1),
                arrayOf(2, 8),
                arrayOf(3, 4)
            )
            println(grid[0][0] + grid[1][1] + grid[2][0])
        }
    "#,
    );
    assert_eq!(out, &["20"]);
}

#[test]
fn test_array_of_chars_and_ascii() {
    let out = run_prints(
        r#"
        fun main() {
            val letters = arrayOf('a', 'b', 'c')
            println(letters[0].toString())
            println(letters[1].code)
            println(letters.size)
        }
    "#,
    );
    assert_eq!(out, &["a", "98", "3"]);
}

#[test]
fn test_array_of_booleans_and_count_true() {
    let out = run_prints(
        r#"
        fun main() {
            val flags = arrayOf(true, false, true, true)
            var trueCount = 0
            for (flag in flags) {
                if (flag) trueCount += 1
            }
            println(trueCount)
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_array_find_by_linear_scan() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(7, 3, 9, 4)
            var found = -1
            var index = 0
            for (value in nums) {
                if (value == 9) {
                    found = value
                    break
                }
                index += 1
            }
            println(found)
            println(index)
        }
    "#,
    );
    assert_eq!(out, &["9", "2"]);
}

#[test]
fn test_array_find_first_false_without_break() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(1, 2, 3, 4)
            var missing = true
            for (value in nums) {
                if (value == 5) {
                    missing = false
                }
            }
            println(missing)
        }
    "#,
    );
    assert_eq!(out, &["true"]);
}

#[test]
fn test_array_range_for_indices_bounds() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(1, 2, 3, 4, 5)
            var evenTotal = 0
            for (i in nums.indices) {
                if (i % 2 == 1) evenTotal += nums[i]
            }
            println(evenTotal)
            println(nums.lastIndex)
        }
    "#,
    );
    assert_eq!(out, &["6", "4"]);
}

#[test]
fn test_array_reverse_manual_copy() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(1, 2, 3, 4)
            val reversed = Array(nums.size) { idx -> nums[nums.size - 1 - idx] }
            var out = ""
            for (value in reversed) {
                out += value.toString()
            }
            println(out)
            println(reversed[0] + reversed[3])
        }
    "#,
    );
    assert_eq!(out, &["4321", "5"]);
}

#[test]
fn test_array_append_like_with_copy_and_set() {
    let out = run_prints(
        r#"
        fun main() {
            val base = arrayOf(1, 2)
            val extended = Array(base.size + 1) { index ->
                if (index < base.size) base[index] else 3
            }
            println(extended.size)
            println(extended[2])
            println(extended[0] + extended[1] + extended[2])
        }
    "#,
    );
    assert_eq!(out, &["3", "3", "6"]);
}

#[test]
fn test_array_with_homogeneous_structures() {
    let out = run_prints(
        r#"
        interface Item {
            fun value(): Int
        }

        class NumberItem(val v: Int) : Item {
            override fun value(): Int = v
        }

        fun main() {
            val boxed: Array<Item> = arrayOf(NumberItem(1), NumberItem(2), NumberItem(3))
            var total = 0
            for (item in boxed) {
                total += item.value()
            }
            println(total)
            println(boxed[1].value())
        }
    "#,
    );
    assert_eq!(out, &["6", "2"]);
}

#[test]
fn test_array_of_pairs_accumulation() {
    let out = run_prints(
        r#"
        fun main() {
            val pairs = arrayOf(Pair(1, 3), Pair(2, 4), Pair(5, 6))
            var sum = 0
            for (item in pairs) {
                sum += item.first + item.second
            }
            println(sum)
            val head = pairs[0]
            println(head.first + head.second)
        }
    "#,
    );
    assert_eq!(out, &["21", "4"]);
}

#[test]
fn test_array_of_functions_dispatch() {
    let out = run_prints(
        r#"
        fun main() {
            val ops = arrayOf(
                { x: Int -> x + 1 },
                { x: Int -> x * 2 },
                { x: Int -> x * x }
            )
            println(ops[0](3))
            println(ops[1](4))
            println(ops[2](5))
        }
    "#,
    );
    assert_eq!(out, &["4", "8", "25"]);
}

#[test]
fn test_array_last_index_mutation_pattern() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(2, 4, 6, 8)
            nums[nums.lastIndex] = nums[0] + nums[1]
            var output = 0
            for (i in 0 until nums.size) {
                output += nums[i]
            }
            println(nums[3])
            println(output)
        }
    "#,
    );
    assert_eq!(out, &["6", "18"]);
}

#[test]
fn test_array_zero_initialization_with_repeat() {
    let out = run_prints(
        r#"
        fun main() {
            val zeros = Array(4) { 0 }
            println(zeros.size)
            var total = 0
            for (value in zeros) {
                total += value
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["4", "0"]);
}

#[test]
fn test_array_while_sum_until_stop() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(3, 1, 0, 9, 2)
            var index = 0
            var sum = 0
            while (index < nums.size) {
                if (nums[index] == 0) {
                    break
                }
                sum += nums[index]
                index += 1
            }
            println(sum)
            println(index)
        }
    "#,
    );
    assert_eq!(out, &["4", "2"]);
}

#[test]
fn test_array_nested_loops_early_continue() {
    let out = run_prints(
        r#"
        fun main() {
            val grid = arrayOf(
                arrayOf(1, 2, 3),
                arrayOf(4, 0, 6),
                arrayOf(7, 8, 9)
            )
            var total = 0
            for (row in grid) {
                for (value in row) {
                    if (value == 0) {
                        continue
                    }
                    total += value
                }
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["40"]);
}

#[test]
fn test_array_singleton_and_size_and_index() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(99)
            println(nums.size)
            println(nums[0])
            println(nums.indices)
        }
    "#,
    );
    assert_eq!(out, &["1", "99", "0..0"]);
}

#[test]
fn test_array_range_filter_without_library() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(1, 2, 3, 4, 5, 6)
            var odds = ""
            for (i in nums.indices) {
                if (i % 2 == 1) {
                    odds += nums[i].toString()
                }
            }
            println(odds)
        }
    "#,
    );
    assert_eq!(out, &["26"]);
}

#[test]
fn test_array_plus_operator_is_not_mutating_source() {
    let out = run_prints(
        r#"
        fun main() {
            val left = arrayOf(1, 2)
            val right = arrayOf(3, 4)
            val joined = left + right
            println(joined.joinToString(","))
            left[0] = 9
            println(joined[0])
        }
    "#,
    );
    assert_eq!(out, &["1,2,3,4", "1"]);
}

#[test]
fn test_array_to_mutable_list_and_mutation_contract() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(1, 2, 3)
            val mutable = nums.toMutableList()
            mutable.add(4)
            mutable[0] = 9
            println(nums.joinToString(","))
            println(mutable.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "9,2,3,4"]);
}

#[test]
fn test_array_content_deep_hashcode_is_stable() {
    let out = run_prints(
        r#"
        fun main() {
            val nested = arrayOf(arrayOf(1, 2), arrayOf(3, 4))
            println(nested.contentDeepHashCode() > 0)
            println(nested.contentDeepToString())
        }
    "#,
    );
    assert_eq!(out, &["true", "[[1, 2], [3, 4]]"]);
}

#[test]
fn test_array_first_last_and_singleton_helpers() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(8, 6, 7)
            println(nums.first())
            println(nums.last())
            println(arrayOf("solo").single())
            println(nums.take(2).joinToString(","))
            println(nums.drop(1).joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["8", "7", "solo", "8,6", "6,7"]);
}

#[test]
fn test_array_count_and_any_all_none() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(1, 3, 5, 6)
            println(nums.count { it > 3 })
            println(nums.any { it == 6 })
            println(nums.all { it > 0 })
            println(nums.none { it < 0 })
        }
    "#,
    );
    assert_eq!(out, &["2", "true", "true", "true"]);
}

#[test]
fn test_array_find_last_if_empty() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(2, 4, 6)
            println(nums.find { it > 5 })
            println(nums.findLast { it < 0 } ?: -1)
            println(nums.firstOrNull { it == 10 } ?: "missing")
        }
    "#,
    );
    assert_eq!(out, &["6", "-1", "missing"]);
}

#[test]
fn test_array_slice_with_ranges() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(10, 20, 30, 40, 50)
            val part = nums.sliceArray(1..3)
            val tail = nums.copyOfRange(3, 5)
            println(part.joinToString(","))
            println(tail.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["20,30,40", "40,50"]);
}

#[test]
fn test_array_filter_map_reduce_pipeline() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(1, 2, 3, 4, 5)
            val result = nums.filter { it % 2 == 1 }
                .map { it * 2 }
                .reduce { acc, value -> acc + value }
            println(result)
        }
    "#,
    );
    assert_eq!(out, &["18"]);
}

#[test]
fn test_array_grouping_by_predicate() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(1, 2, 3, 4, 5, 6)
            val grouped = nums.groupBy { if (it % 2 == 0) "even" else "odd" }
            val even = grouped["even"] ?: arrayOf()
            val odd = grouped["odd"] ?: arrayOf()
            println(even.joinToString(","))
            println(odd.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,4,6", "1,3,5"]);
}

#[test]
fn test_array_with_map_indexed_side_effects() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf("a", "b", "c")
            var marker = ""
            nums.mapIndexed { index, value ->
                marker += index.toString() + value
            }
            println(marker)
        }
    "#,
    );
    assert_eq!(out, &["0a1b2c"]);
}

#[test]
fn test_nested_array_flat_map_to_depth_one() {
    let out = run_prints(
        r#"
        fun main() {
            val buckets = arrayOf(arrayOf(1, 2), arrayOf(3), arrayOf(4, 5))
            val flattened = buckets.flatMap { it.toList() }.toTypedArray()
            println(flattened.joinToString(","))
            println(flattened.size)
        }
    "#,
    );
    assert_eq!(out, &["1,2,3,4,5", "5"]);
}

#[test]
fn test_array_copy_of_throws_on_invalid_range() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = arrayOf(1, 2, 3)
            try {
                nums.copyOfRange(3, 2)
            } catch (e: IllegalArgumentException) {
                println("bad")
            }
        }
    "#,
    );
    assert_eq!(out, &["bad"]);
}

#[test]
fn test_array_fill_and_fill_range_mutate_existing_values() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = IntArray(5) { it }
            nums.fill(9)
            println(nums.joinToString(","))
            val src = intArrayOf(1, 2, 3, 4, 5)
            src.fill(7, 1, 4)
            println(src.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["9,9,9,9,9", "1,7,7,7,5"]);
}

#[test]
fn test_array_to_list_projection_is_snapshot_for_references() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = IntArray(3) { it + 1 }
            val snapshot = nums.toList()
            nums[0] = 9
            println(snapshot.joinToString(","))
            println(nums.joinToString(","))
            println(snapshot[1])
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "9,2,3", "2"]);
}

#[test]
fn test_array_join_to_string_formats_empty_and_nested() {
    let out = run_prints(
        r#"
        fun main() {
            val none = arrayOf<Int>()
            val nested = arrayOf(arrayOf(1), arrayOf(2, 3))
            println(none.joinToString(","))
            println(nested.contentDeepToString())
            println(arrayOf("a").contentToString())
        }
    "#,
    );
    assert_eq!(out, &["", "[[1], [2, 3]]", "[a]"]);
}

#[test]
fn test_array_content_equals_distinguishes_reference_identity() {
    let out = run_prints(
        r#"
        fun main() {
            val left = arrayOf(arrayOf(1), arrayOf(2))
            val same = arrayOf(arrayOf(1), arrayOf(2))
            val deepA = left.contentDeepEquals(same)
            val sameRef = left.contentEquals(same)
            println(deepA)
            println(sameRef)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}
