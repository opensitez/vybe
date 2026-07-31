use crate::helpers::run_prints;

#[test]
fn test_empty_set_basics() {
    let out = run_prints(r#"
        fun main() {
            val values = emptySet<Int>()
            println(values.isEmpty())
            println(values.size)
            println(values.contains(1))
        }
    "#);
    assert_eq!(out, &["true", "0", "false"]);
}

#[test]
fn test_set_identity_with_duplicates() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(1, 1, 2, 2, 3)
            println(values.size)
            println(values.contains(2))
            println(values.contains(4))
        }
    "#);
    assert_eq!(out, &["3", "true", "false"]);
}

#[test]
fn test_mutable_set_add_and_remove() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableSetOf(1, 2)
            println(values.add(3))
            println(values.remove(2))
            println(values.size)
            println(values.contains(2))
        }
    "#);
    assert_eq!(out, &["true", "true", "2", "false"]);
}

#[test]
fn test_set_add_duplicate_is_false() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableSetOf(1, 2)
            println(values.add(2))
            println(values.size)
        }
    "#);
    assert_eq!(out, &["false", "2"]);
}

#[test]
fn test_set_union_union_operator() {
    let out = run_prints(r#"
        fun main() {
            val a = setOf(1, 2, 3)
            val b = setOf(3, 4, 5)
            val merged = a union b
            println(merged.size)
            println(merged.contains(4))
            println(merged.contains(1))
        }
    "#);
    assert_eq!(out, &["5", "true", "true"]);
}

#[test]
fn test_set_union_operator_plus() {
    let out = run_prints(r#"
        fun main() {
            val a = setOf(1, 2)
            val b = setOf(2, 3)
            val merged = a + b
            println(merged.size)
            println(merged.contains(3))
            println((a + b).contains(1))
        }
    "#);
    assert_eq!(out, &["3", "true", "true"]);
}

#[test]
fn test_set_intersection() {
    let out = run_prints(r#"
        fun main() {
            val a = setOf(1, 2, 3)
            val b = setOf(2, 4, 3)
            val overlap = a intersect b
            println(overlap.size)
            println(overlap.contains(2))
            println(overlap.contains(4))
        }
    "#);
    assert_eq!(out, &["2", "true", "false"]);
}

#[test]
fn test_set_difference() {
    let out = run_prints(r#"
        fun main() {
            val a = setOf(1, 2, 3)
            val b = setOf(2, 4)
            val remaining = a - b
            println(remaining.size)
            println(remaining.contains(2))
            println(remaining.contains(1))
        }
    "#);
    assert_eq!(out, &["2", "false", "true"]);
}

#[test]
fn test_set_retain_all_operation() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3, 4)
            println(values.retainAll(setOf(2, 4)))
            println(values.size)
            println(values.contains(1))
            println(values.contains(4))
        }
    "#);
    assert_eq!(out, &["true", "2", "false", "true"]);
}

#[test]
fn test_set_remove_all_operation() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3, 4)
            println(values.removeAll(setOf(2, 4, 6)))
            println(values.size)
            println(values.contains(2))
            println(values.contains(3))
        }
    "#);
    assert_eq!(out, &["true", "2", "false", "true"]);
}

#[test]
fn test_set_clear_and_repopulate() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableSetOf(1, 2)
            values.clear()
            values.add(9)
            values.add(9)
            println(values.size)
            println(values.contains(9))
        }
    "#);
    assert_eq!(out, &["1", "true"]);
}

#[test]
fn test_set_filter_projection() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(1, 2, 3, 4, 5)
            val evens = values.filter { it % 2 == 0 }
            println(evens.size)
            println(evens[1])
        }
    "#);
    assert_eq!(out, &["2", "4"]);
}

#[test]
fn test_set_all_any_none() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(2, 4, 6)
            println(values.all { it % 2 == 0 })
            println(values.any { it > 5 })
            println(values.none { it < 0 })
            println(values.any { it > 10 })
        }
    "#);
    assert_eq!(out, &["true", "true", "true", "false"]);
}

#[test]
fn test_set_sum_and_avg_via_fold() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(1, 2, 3, 4)
            val total = values.fold(0) { acc, value -> acc + value }
            val avg = total / values.size
            println(total)
            println(avg)
        }
    "#);
    assert_eq!(out, &["10", "2"]);
}

#[test]
fn test_set_iteration_order_and_aggregation() {
    let out = run_prints(r#"
        fun main() {
            val values = linkedSetOf(3, 1, 2, 4)
            var output = ""
            for (value in values) {
                output += value.toString()
            }
            println(output)
            println(values.first())
            println(values.elementAt(2))
        }
    "#);
    assert_eq!(out, &["3124", "3", "2"]);
}

#[test]
fn test_to_mutable_set_roundtrip() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(1, 2, 3)
            val mutable = values.toMutableSet()
            mutable.remove(2)
            mutable.add(4)
            println(mutable.size)
            println(mutable.contains(2))
            println(mutable.contains(4))
        }
    "#);
    assert_eq!(out, &["3", "false", "true"]);
}

#[test]
fn test_to_set_from_list_preserves_uniqueness() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(1, 2, 2, 3, 3, 3)
            val unique = values.toSet()
            println(unique.size)
            println(unique.contains(3))
            println(unique.toString())
        }
    "#);
    assert_eq!(out, &["3", "true", "[1, 2, 3]"]);
}

#[test]
fn test_hashset_vs_linkedhashset_mutability() {
    let out = run_prints(r#"
        fun main() {
            val linked = linkedSetOf(1, 2, 3)
            val hash = hashSetOf(3, 2, 1)
            println(linked.size)
            println(hash.size)
            println((linked - setOf(2)).contains(2))
            println((hash + setOf(4)).contains(4))
        }
    "#);
    assert_eq!(out, &["3", "3", "false", "true"]);
}

#[test]
fn test_empty_set_intersection() {
    let out = run_prints(r#"
        fun main() {
            val left = emptySet<Int>()
            val right = setOf(1, 2)
            val result = left.intersect(right)
            println(result.isEmpty())
            println(result.size)
        }
    "#);
    assert_eq!(out, &["true", "0"]);
}

#[test]
fn test_set_of_nested_equality() {
    let out = run_prints(r#"
        fun main() {
            val groups = setOf(setOf(1, 2), setOf(1, 2), setOf(2, 1))
            println(groups.size)
            println(groups.contains(setOf(2, 1)))
        }
    "#);
    assert_eq!(out, &["1", "true"]);
}

#[test]
fn test_set_with_custom_data_class_equality() {
    let out = run_prints(r#"
        data class PairKey(val left: Int, val right: Int)

        fun main() {
            val values = setOf(PairKey(1, 2), PairKey(1, 2), PairKey(2, 1))
            println(values.size)
            println(values.contains(PairKey(1, 2)))
        }
    "#);
    assert_eq!(out, &["2", "true"]);
}

#[test]
fn test_set_with_string_predicate() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf("alpha", "beta", "gamma")
            println(values.count { it.length >= 5 })
            println(values.joinToString("|"))
            println(values.any { it.startsWith("a") })
        }
    "#);
    assert_eq!(out, &["2", "alpha|beta|gamma", "true"]);
}

#[test]
fn test_set_mapping_changes_type() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(1, 2, 3, 4)
            val mapped = values.map { it * 2 }
            val restored = mapped.toSet()
            println(mapped.size)
            println(restored.size)
            println(restored.contains(8))
        }
    "#);
    assert_eq!(out, &["4", "4", "true"]);
}

#[test]
fn test_sorted_set_view() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(4, 2, 1, 3)
            val sorted = values.toSortedSet()
            println(sorted.first())
            println(sorted.last())
            println(sorted.size)
        }
    "#);
    assert_eq!(out, &["1", "4", "4"]);
}

#[test]
fn test_set_min_max_aggregation() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(5, 1, 9, 3)
            println(values.minOrNull())
            println(values.maxOrNull())
        }
    "#);
    assert_eq!(out, &["1", "9"]);
}

#[test]
fn test_set_contains_all_predicate() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(1, 2, 3, 4)
            println(values.containsAll(listOf(1, 4)))
            println(values.containsAll(listOf(1, 6)))
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_set_mutable_remove_if_supported() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3, 4)
            println(values.removeAll { it % 2 == 0 })
            println(values.size)
            println(values.contains(2))
            println(values.contains(3))
        }
    "#);
    assert_eq!(out, &["true", "2", "false", "true"]);
}

#[test]
fn test_set_shallow_copy_and_distinct_reference() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3)
            val copied = values.toMutableSet()
            copied.add(4)
            values.remove(1)
            println(values.size)
            println(copied.size)
            println(values.contains(1))
            println(copied.contains(1))
        }
    "#);
    assert_eq!(out, &["2", "4", "false", "true"]);
}

#[test]
fn test_set_union_with_empty() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(1, 2)
            val merged = values + emptySet<Int>()
            println(merged.size)
            println(merged == values)
            println(emptySet<Int>() + values == values)
        }
    "#);
    assert_eq!(out, &["2", "true", "true"]);
}

#[test]
fn test_set_filter_empty_result() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(1, 2, 3)
            val none = values.filter { it > 10 }.toSet()
            println(none.isEmpty())
            println(none.size)
        }
    "#);
    assert_eq!(out, &["true", "0"]);
}

#[test]
fn test_set_take_drop_analogue() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(1, 2, 3, 4, 5)
            val firstTwo = values.take(2)
            val dropped = values.drop(2)
            println(firstTwo.size)
            println(dropped.size)
            println(firstTwo.contains(1))
            println(dropped.contains(5))
        }
    "#);
    assert_eq!(out, &["2", "3", "true", "true"]);
}

#[test]
fn test_set_join_and_joined_string() {
    let out = run_prints(r#"
        fun main() {
            val values = linkedSetOf(1, 2, 3)
            println(values.joinToString(","))
            println(values.joinToString("|") { it.toString() })
            println(values.joinToString("") { (it * 2).toString() })
        }
    "#);
    assert_eq!(out, &["1,2,3", "1|2|3", "246"]);
}

#[test]
fn test_set_retain_all_with_empty_collection() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3)
            println(values.retainAll(emptySet<Int>()))
            println(values.isEmpty())
            println(values.size)
        }
    "#);
    assert_eq!(out, &["true", "true", "0"]);
}

#[test]
fn test_set_subtract_operator() {
    let out = run_prints(r#"
        fun main() {
            val left = setOf(1, 2, 3, 4)
            val right = setOf(2, 4, 6)
            val remaining = left - right
            println(remaining.size)
            println(remaining.contains(1))
            println(remaining.contains(2))
            println(remaining.contains(6))
        }
    "#);
    assert_eq!(out, &["2", "true", "false", "false"]);
}

#[test]
fn test_set_plus_with_list_input_is_union_like() {
    let out = run_prints(r#"
        fun main() {
            val source = setOf(1, 2)
            val merged = source + listOf(2, 3, 4)
            println(merged.size)
            println(merged.contains(3))
            println(merged.contains(1))
        }
    "#);
    assert_eq!(out, &["4", "true", "true"]);
}

#[test]
fn test_set_union_intersection_do_not_mutate_operands() {
    let out = run_prints(r#"
        fun main() {
            val a = linkedSetOf(1, 2, 3)
            val b = setOf(3, 4)
            val union = a union b
            val inter = a intersect b
            println(union.size)
            println(inter.size)
            println(a.size)
            println(b.size)
            println(a.contains(4))
            println(b.contains(1))
        }
    "#);
    assert_eq!(out, &["4", "1", "3", "2", "false", "false"]);
}

#[test]
fn test_set_minmax_or_null_empty_cases() {
    let out = run_prints(r#"
        fun main() {
            println(setOf<Int>().minOrNull() ?: -1)
            println(setOf<Int>().maxOrNull() ?: -1)
        }
    "#);
    assert_eq!(out, &["-1", "-1"]);
}

#[test]
fn test_set_ordered_iteration_on_linked_set_preserved_after_mutation() {
    let out = run_prints(r#"
        fun main() {
            val values = linkedSetOf(2, 1)
            values.add(3)
            values.remove(1)
            values.add(1)
            var order = ""
            for (value in values) {
                order += value.toString()
            }
            println(order)
            println(values.size)
        }
    "#);
    assert_eq!(out, &["231", "3"]);
}

#[test]
fn test_set_to_list_and_array_roundtrip_consistent_size() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(1, 2, 3)
            val list = values.toList()
            val array = values.toTypedArray()
            println(list.size)
            println(array.size)
            println(list.contains(2))
            println(array.size == list.size)
        }
    "#);
    assert_eq!(out, &["3", "3", "true", "true"]);
}

#[test]
fn test_set_iterator_next_after_end_throws() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(1)
            val it = values.iterator()
            println(it.next())
            try {
                it.next()
            } catch (e: NoSuchElementException) {
                println("done")
            }
        }
    "#);
    assert_eq!(out, &["1", "done"]);
}

#[test]
fn test_set_remove_and_contains_all_edges() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3)
            println(values.remove(99))
            println(values.remove(2))
            println(values.containsAll(setOf(1, 2, 3)))
            println(values.containsAll(setOf(1, 3)))
        }
    "#);
    assert_eq!(out, &["false", "true", "false", "true"]);
}

#[test]
fn test_set_partition_and_counts() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(1, 2, 3, 4, 5)
            val (small, large) = values.partition { it < 4 }
            println(small.joinToString(","))
            println(large.joinToString(","))
            println(small.size + large.size)
        }
    "#);
    assert_eq!(out, &["1,2,3", "4,5", "5"]);
}

#[test]
fn test_set_with_sequence_operations() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(1, 2, 3, 4, 5)
            val sequence = values.asSequence()
            val total = sequence.filter { it > 2 }.sum()
            println(total)
        }
    "#);
    assert_eq!(out, &["12"]);
}

#[test]
fn test_set_fold_with_identity() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(1, 2, 3)
            println(values.fold(0) { acc, item -> acc + item })
            println(values.reduce { acc, item -> acc * item })
        }
    "#);
    assert_eq!(out, &["6", "6"]);
}

#[test]
fn test_set_average_on_empty_handled_by_exception() {
    let out = run_prints(r#"
        fun main() {
            val avg = setOf<Int>().average()
            println(avg.isNaN())
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_set_flatten_nested_set_depth_one() {
    let out = run_prints(r#"
        fun main() {
            val nested: Set<Set<Int>> = setOf(setOf(1, 2), setOf(3, 4))
            val flattened = nested.flatten().toSet()
            println(flattened.size)
            println(flattened.contains(3))
        }
    "#);
    assert_eq!(out, &["4", "true"]);
}

#[test]
fn test_set_zip_with_index_like_build() {
    let out = run_prints(r#"
        fun main() {
            val values = linkedSetOf("a", "b", "c")
            val zip = values.zipWithNext()
            println(zip.size)
            println(zip[0].first)
            println(zip[0].second)
        }
    "#);
    assert_eq!(out, &["2", "a", "b"]);
}

#[test]
fn test_set_take_while_skip_while_sequence_operators() {
    let out = run_prints(r#"
        fun main() {
            val values = linkedSetOf(1, 2, 3, 4, 5)
            println(values.takeWhile { it < 4 })
            println(values.dropWhile { it < 4 })
        }
    "#);
    assert_eq!(out, &["[1, 2, 3]", "[4, 5]"]);
}

#[test]
fn test_set_lookup_with_mutated_element_hash_breaks() {
    let out = run_prints(r#"
        fun main() {
            data class Item(var id: Int)
            val item = Item(1)
            val values = hashSetOf(item)
            println(values.contains(item))
            item.id = 99
            println(values.contains(item))
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_set_equals_by_value_on_data_class_instances() {
    let out = run_prints(r#"
        fun main() {
            data class Label(val id: Int, val name: String)
            val values = setOf(Label(1, "x"))
            println(values.contains(Label(1, "x")))
            println(values.contains(Label(1, "y")))
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_set_reference_equality_with_plain_class() {
    let out = run_prints(r#"
        fun main() {
            class Label(val id: Int)
            val values = setOf(Label(1))
            println(values.contains(Label(1)))
        }
    "#);
    assert_eq!(out, &["false"]);
}

#[test]
fn test_set_replace_element_after_remove_by_equal_shape() {
    let out = run_prints(r#"
        data class Box(val id: Int)

        fun main() {
            val values = mutableSetOf(Box(1), Box(2))
            values.remove(Box(1))
            values.add(Box(3))
            println(values.size)
            println(values.contains(Box(2)))
            println(values.contains(Box(1)))
            println(values.contains(Box(3)))
        }
    "#);
    assert_eq!(out, &["2", "true", "false", "true"]);
}

#[test]
fn test_set_retain_all_with_self_is_noop() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3)
            println(values.retainAll(values.toSet()))
            println(values.size)
            println(values.contains(1))
            println(values.contains(3))
        }
    "#);
    assert_eq!(out, &["false", "3", "true", "true"]);
}

#[test]
fn test_set_contains_and_size_with_null_element() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf<String?>(null, "first", null)
            println(values.size)
            println(values.contains(null))
            println(values.contains("missing"))
        }
    "#);
    assert_eq!(out, &["2", "true", "false"]);
}

#[test]
fn test_set_retain_all_no_change_returns_false() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3)
            println(values.retainAll(setOf(1, 2, 3)))
            println(values.size)
        }
    "#);
    assert_eq!(out, &["false", "3"]);
}

#[test]
fn test_set_remove_all_with_empty_collection_no_change() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3)
            println(values.removeAll(emptySet<Int>()))
            println(values.size)
        }
    "#);
    assert_eq!(out, &["false", "3"]);
}

#[test]
fn test_set_difference_with_empty_is_copy_like() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(1, 2, 3)
            val remaining = values - emptySet<Int>()
            println(remaining.size)
            println(remaining == values)
        }
    "#);
    assert_eq!(out, &["3", "true"]);
}

#[test]
fn test_sorted_set_with_string_ordering() {
    let out = run_prints(r#"
        fun main() {
            val ordered = sortedSetOf(3, 1, 2)
            println(ordered.first())
            println(ordered.last())
            println(ordered.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1", "3", "1,2,3"]);
}

#[test]
fn test_mutable_set_plus_assign_with_elements() {
    let out = run_prints(r#"
        fun main() {
            val base = mutableSetOf(1, 2)
            val snapshot = base.toSet()
            base += setOf(2, 3, 4)
            println(base.size)
            println(snapshot.size)
            println(base.contains(4))
        }
    "#);
    assert_eq!(out, &["4", "2", "true"]);
}

#[test]
fn test_mutable_set_minus_assign_with_elements() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3, 4)
            values -= setOf(2, 9, 4)
            println(values.size)
            println(values.contains(2))
            println(values.contains(4))
            println(values.contains(3))
        }
    "#);
    assert_eq!(out, &["2", "false", "false", "true"]);
}

#[test]
fn test_set_add_all_is_snapshot_preserved_for_list() {
    let out = run_prints(r#"
        fun main() {
            val source = setOf(1, 2, 3)
            val copied = source.toMutableSet()
            copied.addAll(listOf(3, 4, 5))
            println(source.toString())
            println(copied.toString())
            println(source.contains(5))
            println(copied.contains(5))
        }
    "#);
    assert_eq!(out, &["[1, 2, 3]", "[1, 2, 3, 4, 5]", "false", "true"]);
}

#[test]
fn test_set_with_nullable_values_distinguishes_null_presence() {
    let out = run_prints(r#"
        fun main() {
            val values: Set<String?> = setOf("a", null, "b", null)
            println(values.size)
            println(values.contains(null))
            println(values.contains("c"))
        }
    "#);
    assert_eq!(out, &["3", "true", "false"]);
}

#[test]
fn test_set_retain_all_no_change_returns_false_when_unchanged() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3)
            println(values.retainAll(setOf(1, 2, 3)))
            println(values.size)
            println(values.toString())
        }
    "#);
    assert_eq!(out, &["false", "3", "[1, 2, 3]"]);
}
