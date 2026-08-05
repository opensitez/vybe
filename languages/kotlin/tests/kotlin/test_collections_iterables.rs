use crate::helpers::run_prints;

#[test]
fn test_list_map_transforms_each_element() {
    let out = run_prints(
        r#"
    fun main() {
            val nums = listOf(1, 2, 3)
            val doubled = nums.map { it * 2 }
            println(doubled.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,4,6"]);
}

#[test]
fn test_list_filter_keeps_matching_predicate() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 2, 3, 4)
            val evens = nums.filter { it % 2 == 0 }
            println(evens.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,4"]);
}

#[test]
fn test_list_filter_not_includes_only_failures() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 2, 3, 4)
            val odds = nums.filterNot { it % 2 == 0 }
            println(odds.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,3"]);
}

#[test]
fn test_list_any_and_all_and_none() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 2, 3, 4)
            println(nums.any { it > 3 })
            println(nums.all { it < 10 })
            println(nums.none { it == 9 })
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_list_count_with_predicate() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 2, 3, 4, 5)
            println(nums.count { it % 2 == 1 })
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_list_find_first_matching() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(5, 12, 13, 20)
            println(nums.find { it % 2 == 0 })
            println(nums.findLast { it % 2 == 1 })
        }
    "#,
    );
    assert_eq!(out, &["12", "13"]);
}

#[test]
fn test_list_first_or_null_and_last_or_null() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(9, 1, 5)
            println(nums.firstOrNull { it > 10 } ?: "none")
            println(nums.lastOrNull { it < 3 } ?: "none")
        }
    "#,
    );
    assert_eq!(out, &["none", "1"]);
}

#[test]
fn test_list_single_and_single_or_null() {
    let out = run_prints(
        r#"
        fun main() {
            val only = listOf(42)
            println(only.single())
            println(listOf<Int>().singleOrNull() ?: -1)
        }
    "#,
    );
    assert_eq!(out, &["42", "-1"]);
}

#[test]
fn test_list_fold_left_and_reduce_right() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 2, 3, 4)
            val sum = nums.fold(0) { acc, v -> acc + v }
            val diff = nums.reduceRight { a, b -> a - b }
            println(sum)
            println(diff)
        }
    "#,
    );
    assert_eq!(out, &["10", "-2"]);
}

#[test]
fn test_list_sorted_with_comparator() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf("delta", "a", "charlie", "bb")
            val sorted = nums.sortedWith(compareBy { it.length })
            println(sorted.joinToString(","))
            val reverse = nums.sortedWith(compareByDescending { it.length })
            println(reverse.joinToString(","))
        }
    "#,
    );
    // By LENGTH: a=1, bb=2, delta=5, charlie=7 — so ascending puts delta
    // before charlie, descending the reverse (real Kotlin agrees).
    assert_eq!(out, &["a,bb,delta,charlie", "charlie,delta,bb,a"]);
}

#[test]
fn test_list_sorted_by_and_reversed() {
    let out = run_prints(
        r#"
        fun main() {
            val letters = listOf("pear", "apple", "kiwi")
            println(letters.sorted().joinToString(","))
            println(letters.reversed().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["apple,kiwi,pear", "kiwi,apple,pear"]);
}

#[test]
fn test_list_distinct_and_union_intersection() {
    let out = run_prints(
        r#"
        fun main() {
            val left = listOf(1, 1, 2, 3, 3)
            val right = listOf(3, 3, 4)
            println(left.distinct().joinToString(","))
            println((left.toSet() intersect right.toSet()).joinToString(","))
            println((left.toSet() union right.toSet()).joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "3", "1,2,3,4"]);
}

#[test]
fn test_list_slice_take_drop() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 2, 3, 4, 5)
            println(nums.take(3).joinToString(","))
            println(nums.drop(2).joinToString(","))
            println(nums.slice(1..3).joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "3,4,5", "2,3,4"]);
}

#[test]
fn test_list_take_while_and_drop_while() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 3, 5, 2, 4, 6, 8)
            println(nums.takeWhile { it < 4 }.joinToString(","))
            println(nums.dropWhile { it < 4 }.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,3", "5,2,4,6,8"]);
}

#[test]
fn test_list_chunked_and_windowed() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 2, 3, 4, 5)
            println(nums.chunked(2).joinToString("|") { it.joinToString("-") })
            println(nums.windowed(2).joinToString("|") { it.joinToString("-") })
        }
    "#,
    );
    assert_eq!(out, &["1-2|3-4|5", "1-2|2-3|3-4|4-5"]);
}

#[test]
fn test_list_partition_predicate() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 2, 3, 4, 5, 6)
            val (evens, odds) = nums.partition { it % 2 == 0 }
            println(evens.joinToString(","))
            println(odds.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,4,6", "1,3,5"]);
}

#[test]
fn test_list_join_to_string_separator_prefix_suffix() {
    let out = run_prints(
        r#"
        fun main() {
            val names = listOf("a", "b", "c")
            println(names.joinToString(prefix = "[", postfix = "]", separator = ","))
        }
    "#,
    );
    assert_eq!(out, &["[a,b,c]"]);
}

#[test]
fn test_list_map_not_empty_and_is_not_empty() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 2)
            val empty = listOf<Int>()
            println(nums.isNotEmpty())
            println(empty.isEmpty())
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_list_drop_last_take_last() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 2, 3, 4, 5)
            println(nums.dropLast(2).joinToString(","))
            println(nums.takeLast(3).joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "3,4,5"]);
}

#[test]
fn test_list_contains_all_and_sublist() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 2, 3, 4, 5)
            println(nums.containsAll(listOf(2, 4)))
            println(nums.subList(1, 4).joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["true", "2,3,4"]);
}

#[test]
fn test_list_zip_with_other_list() {
    let out = run_prints(
        r#"
        fun main() {
            val left = listOf("a", "b", "c")
            val right = listOf(1, 2, 3, 4)
            val pairs = left.zip(right) { l, r -> "$l:$r" }
            println(pairs.joinToString("|"))
        }
    "#,
    );
    assert_eq!(out, &["a:1|b:2|c:3"]);
}

#[test]
fn test_list_flat_map_expands_and_maps() {
    let out = run_prints(
        r#"
        fun main() {
            val groups = listOf(
                listOf(1, 2),
                listOf(3, 4)
            )
            val expanded = groups.flatMap { it }
            println(expanded.joinToString(","))
            val mapped = groups.flatMap { inner -> inner.map { it * 10 } }
            println(mapped.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3,4", "10,20,30,40"]);
}

#[test]
fn test_list_flatten_nested_structure() {
    let out = run_prints(
        r#"
        fun main() {
            val nested = listOf(
                listOf("x", "y"),
                listOf(),
                listOf("z")
            )
            println(nested.flatten().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["x,y,z"]);
}

#[test]
fn test_list_associate_by_and_group_by() {
    let out = run_prints(
        r#"
        fun main() {
            val people = listOf(
                Pair("alice", "A"),
                Pair("bob", "A"),
                Pair("charlie", "B")
            )
            val byGroup = people.groupBy { it.second }
            println(byGroup["A"]?.size ?: 0)
            val mapByName = people.associateBy { it.first }
            println(mapByName["charlie"]?.second)
        }
    "#,
    );
    assert_eq!(out, &["2", "B"]);
}

#[test]
fn test_list_min_max_or_null_empty() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(5, 2, 9, 1)
            println(nums.minOrNull())
            println(nums.maxOrNull())
            println(listOf<Int>().minOrNull() ?: -1)
        }
    "#,
    );
    assert_eq!(out, &["1", "9", "-1"]);
}

#[test]
fn test_list_min_by_and_max_by() {
    let out = run_prints(
        r#"
        fun main() {
            val words = listOf("pear", "apple", "banana")
            println(words.minByOrNull { it.length } ?: "")
            println(words.maxByOrNull { it.length } ?: "")
        }
    "#,
    );
    assert_eq!(out, &["pear", "banana"]);
}

#[test]
fn test_list_sum_and_average_of_ints() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 2, 3, 4)
            println(nums.sum())
            println(nums.average())
        }
    "#,
    );
    assert_eq!(out, &["10", "2.5"]);
}

#[test]
fn test_list_on_empty_collection_error_paths() {
    let out = run_prints(
        r#"
        fun main() {
            try {
                val x = emptyList<Int>().first()
                println(x)
            } catch (e: NoSuchElementException) {
                println("first_error")
            }
            try {
                val y = emptyList<Int>().elementAt(1)
                println(y)
            } catch (e: IndexOutOfBoundsException) {
                println("element_error")
            }
        }
    "#,
    );
    assert_eq!(out, &["first_error", "element_error"]);
}

#[test]
fn test_element_at_and_element_at_or_null() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(10, 20, 30)
            println(nums.elementAt(1))
            println(nums.elementAtOrNull(5) ?: -1)
        }
    "#,
    );
    assert_eq!(out, &["20", "-1"]);
}

#[test]
fn test_binary_search_and_binary_search_not_found_position() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 3, 5, 7, 9)
            println(nums.binarySearch(5))
            println(nums.binarySearch(6))
        }
    "#,
    );
    // `binarySearch` answers `-(insertion point) - 1` when absent; 6 slots in
    // at index 3 of [1, 3, 5, 7, 9], so the result is -4 (real Kotlin agrees).
    assert_eq!(out, &["2", "-4"]);
}

#[test]
fn test_list_map_indexed_applies_index_and_value() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(10, 20, 30)
            val withIndex = nums.mapIndexed { index, value -> value + index }
            println(withIndex.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["10,21,32"]);
}

#[test]
fn test_list_map_indexed_not_null_filters_by_indexed_predicate() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 2, 3, 4)
            val picked = nums.mapIndexedNotNull { index, value ->
                if (index % 2 == 0) value * 10 else null
            }
            println(picked.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["10,30"]);
}

#[test]
fn test_list_with_index_iteration_binding() {
    let out = run_prints(
        r#"
        fun main() {
            val letters = listOf("a", "b", "c")
            var parts = ""
            for ((i, value) in letters.withIndex()) {
                parts += "${i}:${value};"
            }
            println(parts)
        }
    "#,
    );
    assert_eq!(out, &["0:a;1:b;2:c;"]);
}

#[test]
fn test_list_running_fold_accumulates_prefixes() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 2, 3, 4)
            val prefixes = nums.runningFold(0) { acc, value -> acc + value }
            println(prefixes.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["0,1,3,6,10"]);
}

#[test]
fn test_list_running_reduce_aggregate_chain() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 2, 3, 4)
            val running = nums.runningReduce { acc, value -> acc + value }
            println(running.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,3,6,10"]);
}

#[test]
fn test_list_fold_indexed_weighted_sum() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 2, 3, 4)
            val result = nums.foldIndexed(0) { index, acc, value -> acc + index * value }
            println(result)
        }
    "#,
    );
    assert_eq!(out, &["20"]);
}

#[test]
fn test_list_zip_with_next_produces_adjacent_pairs() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 2, 3, 4)
            val pairs = nums.zipWithNext { a, b -> "$a:$b" }
            println(pairs.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1:2,2:3,3:4"]);
}

#[test]
fn test_mutable_list_operations_modify_size_and_order() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = mutableListOf(1, 2, 3)
            nums.add(4)
            nums.removeAt(1)
            nums[0] = 8
            nums.remove(3)
            println(nums.joinToString(","))
            println(nums.size)
        }
    "#,
    );
    assert_eq!(out, &["8,4", "2"]);
}

#[test]
fn test_mutable_list_get_out_of_bounds_is_exception() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = mutableListOf(10)
            try {
                println(nums[3])
            } catch (e: Exception) {
                println("oob")
            }
        }
    "#,
    );
    assert_eq!(out, &["oob"]);
}

#[test]
fn test_iterable_as_sequence_is_lazy_until_terminal() {
    let out = run_prints(
        r#"
        fun main() {
            val source = listOf(1, 2, 3, 4)
            var mappedCount = 0
            val seq = source.asSequence().map {
                mappedCount += 1
                it * 2
            }
            println(mappedCount)
            val firstTwo = seq.take(2).toList().joinToString(",")
            println(mappedCount)
            println(firstTwo)
        }
    "#,
    );
    assert_eq!(out, &["0", "2", "2,4"]);
}

#[test]
fn test_sequence_take_and_take_while_are_short_circuiting() {
    let out = run_prints(
        r#"
        fun main() {
            var mapped = 0
            val seq = sequenceOf(1, 2, 3, 4, 5).map {
                mapped += 1
                it
            }
            val taken = seq.take(3).toList().joinToString(",")
            println(mapped)
            val bounded = sequenceOf(1, 2, 3, 4, 5)
                .map { it }
                .takeWhile { it < 4 }
                .toList()
                .joinToString(",")
            println(bounded)
            println(mapped)
        }
    "#,
    );
    assert_eq!(out, &["3", "1,2,3", "3"]);
}

#[test]
fn test_reversed_list_is_independent_snapshot() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = mutableListOf(1, 2, 3)
            val snapshot = nums.reversed()
            nums[0] = 9
            nums.add(4)
            println(snapshot.joinToString(","))
            println(nums.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["3,2,1", "9,2,3,4"]);
}

#[test]
fn test_as_reversed_is_live_view_of_mutable_list() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = mutableListOf(1, 2, 3)
            val view = nums.asReversed()
            println(view.joinToString(","))
            nums.add(4)
            println(view.joinToString(","))
            view[0] = 10
            println(nums.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["3,2,1", "4,3,2,1", "10,2,3,1,4"]);
}

#[test]
fn test_associate_by_last_duplicate_key_wins() {
    let out = run_prints(
        r#"
        fun main() {
            val entries = listOf("a" to 1, "b" to 2, "a" to 9)
            val map = entries.associateBy({ it.first }) { it.second }
            println(map["a"])
            println(map["b"])
            println(map.size)
        }
    "#,
    );
    assert_eq!(out, &["9", "2", "2"]);
}

#[test]
fn test_distinct_by_projection_keeps_first_per_group() {
    let out = run_prints(
        r#"
        fun main() {
            val words = listOf("alpha", "alpine", "beta", "breeze", "bravo")
            println(words.distinctBy { it[0] }.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["alpha,beta"]);
}

#[test]
fn test_grouping_by_each_count() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 1, 2, 2, 2, 3)
            val counts = nums.groupingBy { it }.eachCount()
            println(counts[1])
            println(counts[2])
            println(counts[3])
        }
    "#,
    );
    assert_eq!(out, &["2", "3", "1"]);
}

#[test]
fn test_for_each_indexed_emits_expected_indexes_and_values() {
    let out = run_prints(
        r#"
        fun main() {
            val items = listOf("x", "y", "z")
            var marker = ""
            items.forEachIndexed { index, value ->
                marker += "${index}${value}"
            }
            println(marker)
        }
    "#,
    );
    assert_eq!(out, &["0x1y2z"]);
}

#[test]
fn test_on_each_returns_same_mutable_collection_reference() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = mutableListOf(1, 2, 3)
            val same = nums.onEach { }
            same[0] = 9
            println(nums[0])
            println(nums === same)
        }
    "#,
    );
    assert_eq!(out, &["9", "true"]);
}

#[test]
fn test_iterator_consumption_and_end_of_data_error() {
    let out = run_prints(
        r#"
        fun main() {
            val iterator = listOf(1, 2).iterator()
            println(iterator.hasNext())
            println(iterator.next())
            println(iterator.hasNext())
            println(iterator.next())
            println(iterator.hasNext())
            try {
                iterator.next()
                println("should_not_happen")
            } catch (e: NoSuchElementException) {
                println("error")
            }
        }
    "#,
    );
    assert_eq!(out, &["true", "1", "true", "2", "false", "error"]);
}

#[test]
fn test_iterator_reuse_on_iterable_is_distinct_instances() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(1, 2, 3)
            val first = values.iterator()
            while (first.hasNext()) {
                println(first.next())
            }
            val second = values.iterator()
            println(second.hasNext())
            println(second.next())
        }
    "#,
    );
    assert_eq!(out, &["1", "2", "3", "true", "1"]);
}

#[test]
fn test_filter_not_null_and_map_not_null_distinguish_none() {
    let out = run_prints(
        r#"
        fun main() {
            val items = listOf(1, null, 2, null, 3)
            println(items.filterNotNull().joinToString(","))
            val transformed = items.mapNotNull { it?.plus(10) }
            println(transformed.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "11,12,13"]);
}

#[test]
fn test_list_of_not_null_omits_empty_inputs() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOfNotNull(null, 7, null, 0, 4)
            println(values.size)
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["3", "7,0,4"]);
}

#[test]
fn test_chunked_with_transform_and_incomplete_tail() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = (1..5).toList()
            println(nums.chunked(2).joinToString("|") { it.joinToString("-") })
            println(nums.chunked(3) { it.sum() }.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1-2|3-4|5", "6,9"]);
}

#[test]
fn test_sequence_is_lazy_until_terminal_operation() {
    let out = run_prints(
        r#"
        fun main() {
            var called = 0
            val source = sequenceOf(1, 2, 3, 4, 5)
            val transformed = source.map {
                called += 1
                it * 2
            }
            println(called)
            val first = transformed.first()
            println(called)
            println(first)
            val rest = transformed.take(2).toList()
            println(rest.joinToString(","))
            println(called)
        }
    "#,
    );
    assert_eq!(out, &["0", "1", "2", "4,6", "3"]);
}
