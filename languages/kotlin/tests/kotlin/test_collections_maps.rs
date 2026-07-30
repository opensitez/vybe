use crate::helpers::run_prints;

#[test]
fn test_list_creation_and_indexing() {
    let out = run_prints(r#"
        fun main() {
            val nums = listOf(10, 20, 30)
            println(nums.size)
            println(nums[0])
            println(nums[2])
        }
    "#);
    assert_eq!(out, &["3", "10", "30"]);
}

#[test]
fn test_list_filter_and_sum() {
    let out = run_prints(r#"
        fun main() {
            val nums = listOf(1, 2, 3, 4, 5)
            val evens = nums.filter { it % 2 == 0 }
            var total = 0
            for (v in evens) {
                total += v
            }
            println(evens.size)
            println(total)
        }
    "#);
    assert_eq!(out, &["2", "6"]);
}

#[test]
fn test_list_map_projection() {
    let out = run_prints(r#"
        fun main() {
            val nums = listOf(1, 2, 3)
            val doubled = nums.map { it * 2 }
            var total = 0
            for (v in doubled) {
                total += v
            }
            println(doubled[0] + doubled[1] + doubled[2])
            println(total)
        }
    "#);
    assert_eq!(out, &["12", "12"]);
}

#[test]
fn test_list_contains_and_index_of() {
    let out = run_prints(r#"
        fun main() {
            val words = listOf("a", "b", "c")
            println(words.contains("b"))
            println(words.indexOf("c"))
            println(words.lastIndex)
        }
    "#);
    assert_eq!(out, &["true", "2", "2"]);
}

#[test]
fn test_mutable_list_add_remove() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableListOf(1, 3)
            values.add(5)
            values.removeAt(1)
            println(values.size)
            println(values[1])
        }
    "#);
    assert_eq!(out, &["2", "5"]);
}

#[test]
fn test_set_uniqueness_and_lookup() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(1, 2, 2, 3, 1)
            println(values.size)
            println(values.contains(2))
            println(values.contains(4))
        }
    "#);
    assert_eq!(out, &["3", "true", "false"]);
}

#[test]
fn test_mutable_set_update() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableSetOf(1, 2)
            values.add(2)
            values.add(3)
            println(values.size)
            values.remove(1)
            println(values.contains(1))
        }
    "#);
    assert_eq!(out, &["3", "false"]);
}

#[test]
fn test_map_basic_put_get() {
    let out = run_prints(r#"
        fun main() {
            val scores = mapOf("alice" to 3, "bob" to 7)
            println(scores["alice"])
            println(scores["bob"])
            println(scores.size)
        }
    "#);
    assert_eq!(out, &["3", "7", "2"]);
}

#[test]
fn test_mutable_map_update() {
    let out = run_prints(r#"
        fun main() {
            val counters = mutableMapOf("a" to 1, "b" to 2)
            counters["a"] = 4
            println(counters["a"])
            counters.remove("b")
            println(counters.containsKey("b"))
            println(counters.size)
        }
    "#);
    assert_eq!(out, &["4", "false", "1"]);
}

#[test]
fn test_map_iteration_keys_values() {
    let out = run_prints(r#"
        fun main() {
            val data = mapOf("x" to 1, "y" to 2)
            var keys = ""
            var sum = 0
            for (entry in data.entries) {
                keys += entry.key
                sum += entry.value
            }
            println(keys)
            println(sum)
        }
    "#);
    assert_eq!(out, &["xy", "3"]);
}

#[test]
fn test_list_empty_properties() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf<Int>()
            println(values.size)
            println(values.isEmpty())
        }
    "#);
    assert_eq!(out, &["0", "true"]);
}

#[test]
fn test_mutable_list_insert_and_remove_at_index() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableListOf(1, 2, 3)
            values.add(1, 9)
            values.removeAt(2)
            println(values.size)
            println(values[1])
            println(values[2])
        }
    "#);
    assert_eq!(out, &["3", "9", "3"]);
}

#[test]
fn test_mutable_list_insert_and_update() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableListOf("a", "c")
            values.add(1, "b")
            values[2] = "d"
            var output = ""
            for (value in values) {
                output += value
            }
            println(output)
            println(values.size)
        }
    "#);
    assert_eq!(out, &["abd", "3"]);
}

#[test]
fn test_mutable_list_remove_value_return_value() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableListOf(1, 2, 3)
            println(values.remove(2))
            println(values.remove(9))
            println(values.size)
        }
    "#);
    assert_eq!(out, &["true", "false", "2"]);
}

#[test]
fn test_list_index_lookup_and_last_position() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(5, 6, 7, 6, 8)
            println(values.indexOf(6))
            println(values.lastIndexOf(6))
            var output = ""
            for (i in values.indices) {
                if (i % 2 == 1) {
                    output += values[i].toString()
                }
            }
            println(values.size)
            println(output)
        }
    "#);
    assert_eq!(out, &["1", "3", "4", "68"]);
}

#[test]
fn test_list_iteration_with_manual_early_exit() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(4, 8, 12, 16)
            var sum = 0
            var found = false
            for (value in values) {
                sum += value
                if (sum > 15) {
                    found = true
                    break
                }
            }
            println(sum)
            println(found)
        }
    "#);
    assert_eq!(out, &["24", "true"]);
}

#[test]
fn test_set_starts_empty_and_adds_unique() {
    let out = run_prints(r#"
        fun main() {
            val ids = mutableSetOf<Int>()
            ids.add(1)
            ids.add(2)
            ids.add(2)
            println(ids.size)
            println(ids.contains(1))
            println(ids.contains(3))
        }
    "#);
    assert_eq!(out, &["2", "true", "false"]);
}

#[test]
fn test_set_remove_and_is_empty() {
    let out = run_prints(r#"
        fun main() {
            val ids = mutableSetOf(1, 2, 3)
            println(ids.remove(2))
            println(ids.remove(4))
            println(ids.isNotEmpty())
            ids.remove(1)
            ids.remove(3)
            println(ids.isEmpty())
            println(ids.size)
        }
    "#);
    assert_eq!(out, &["true", "false", "true", "true", "0"]);
}

#[test]
fn test_set_sum_via_iteration() {
    let out = run_prints(r#"
        fun main() {
            val ids = setOf(3, 4, 5)
            var total = 0
            for (id in ids) {
                total += id
            }
            println(total)
            println(ids.size == 3)
        }
    "#);
    assert_eq!(out, &["12", "true"]);
}

#[test]
fn test_set_contains_value_after_mutation() {
    let out = run_prints(r#"
        fun main() {
            val ids = mutableSetOf(10, 20)
            ids.add(30)
            ids.add(20)
            ids.remove(10)
            println(ids.contains(10))
            println(ids.contains(30))
            println(ids.size)
        }
    "#);
    assert_eq!(out, &["false", "true", "2"]);
}

#[test]
fn test_map_lookup_missing_key_fallback() {
    let out = run_prints(r#"
        fun main() {
            val scores = mapOf("a" to 1, "b" to 2)
            println(scores["missing"] ?: -1)
            println(scores.containsKey("missing"))
            println(scores.get("b") ?: -1)
        }
    "#);
    assert_eq!(out, &["-1", "false", "2"]);
}

#[test]
fn test_map_mutable_put_and_replace() {
    let out = run_prints(r#"
        fun main() {
            val counts = mutableMapOf("a" to 1)
            counts["a"] = 4
            counts["b"] = 2
            println(counts["a"])
            println(counts["b"])
            println(counts.size)
        }
    "#);
    assert_eq!(out, &["4", "2", "2"]);
}

#[test]
fn test_map_remove_and_reinsert() {
    let out = run_prints(r#"
        fun main() {
            val items = mutableMapOf("a" to 1, "b" to 2, "c" to 3)
            println(items.remove("b"))
            println(items.remove("x"))
            println(items.size)
            items["b"] = 9
            println(items["b"])
            println(items.size)
        }
    "#);
    assert_eq!(out, &["2", "null", "2", "9", "3"]);
}

#[test]
fn test_map_entries_aggregation() {
    let out = run_prints(r#"
        fun main() {
            val metrics = mapOf("read" to 5, "write" to 7, "update" to 3)
            var total = 0
            var hasUpdate = false
            for ((name, value) in metrics) {
                total += value
                if (name == "update") {
                    hasUpdate = true
                }
            }
            println(total)
            println(hasUpdate)
            println(metrics.size)
        }
    "#);
    assert_eq!(out, &["15", "true", "3"]);
}

#[test]
fn test_map_nested_values_sum() {
    let out = run_prints(r#"
        fun main() {
            val buckets = mapOf(
                "left" to listOf(1, 2, 3),
                "right" to listOf(4, 5)
            )
            var total = 0
            for (value in buckets["left"]!!) {
                total += value
            }
            for (value in buckets["right"]!!) {
                total += value
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["15"]);
}

#[test]
fn test_map_with_null_values() {
    let out = run_prints(r#"
        fun main() {
            val values = mapOf("x" to null, "y" to 2)
            println(values["x"])
            println(values.containsKey("x"))
            println(values["z"] ?: -1)
        }
    "#);
    assert_eq!(out, &["null", "true", "-1"]);
}

#[test]
fn test_map_size_after_clear() {
    let out = run_prints(r#"
        fun main() {
            val data = mutableMapOf("a" to 1, "b" to 2)
            data.clear()
            println(data.isEmpty())
            data["z"] = 8
            println(data.size)
            println(data["z"])
        }
    "#);
    assert_eq!(out, &["true", "1", "8"]);
}

#[test]
fn test_map_key_membership_across_nested_collections() {
    let out = run_prints(r#"
        fun main() {
            val registry = mapOf(
                "admin" to listOf("read", "write"),
                "guest" to listOf("read")
            )
            var canWrite = false
            val role = "admin"
            for ((user, permissions) in registry) {
                if (user == role) {
                    for (perm in permissions) {
                        if (perm == "write") {
                            canWrite = true
                        }
                    }
                }
            }
            println(canWrite)
            println(registry.containsKey("guest"))
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_map_collect_keys_to_string() {
    let out = run_prints(r#"
        fun main() {
            val data = mapOf("a" to 10, "b" to 20, "c" to 30)
            var keys = ""
            for (k in data.keys) {
                keys += k
            }
            println(keys)
            println(data.keys.size)
        }
    "#);
    assert_eq!(out, &["abc", "3"]);
}

#[test]
fn test_mutable_list_clear_and_reuse() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableListOf(1, 2, 3)
            values.clear()
            println(values.size)
            values.add(4)
            values.add(5)
            println(values.size)
            println(values[0] + values[1])
        }
    "#);
    assert_eq!(out, &["0", "2", "9"]);
}

#[test]
fn test_map_duplicate_keys_keep_the_last_value() {
    let out = run_prints(r#"
        fun main() {
            val scores = mapOf("a" to 1, "b" to 2, "a" to 7, "b" to 9)
            println(scores.size)
            println(scores["a"])
            println(scores["b"])
        }
    "#);
    assert_eq!(out, &["2", "7", "9"]);
}

#[test]
fn test_mutable_map_update_does_not_reorder_existing_keys() {
    let out = run_prints(r#"
        fun main() {
            val state = linkedMapOf("first" to 1, "second" to 2)
            state["first"] = 9
            state["first"] = 11
            var keys = ""
            for ((key, _) in state) {
                keys += key
            }
            println(keys)
            println(state["first"])
            println(state.size)
        }
    "#);
    assert_eq!(out, &["firstsecond", "11", "2"]);
}

#[test]
fn test_map_key_view_reflects_mutations() {
    let out = run_prints(r#"
        fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            val keys = map.keys
            println(keys.contains("a"))
            map["c"] = 3
            println(keys.size)
            map.remove("a")
            println(keys.contains("a"))
            println(keys.contains("c"))
            map.clear()
            println(keys.isEmpty())
        }
    "#);
    assert_eq!(out, &["true", "3", "false", "true", "true"]);
}

#[test]
fn test_map_value_view_tracks_updates() {
    let out = run_prints(r#"
        fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            val values = map.values
            map["a"] = 4
            map["c"] = 3
            var sum = 0
            for (value in values) {
                sum += value
            }
            println(sum)
            map.remove("b")
            println(values.size)
            println(map.size)
        }
    "#);
    assert_eq!(out, &["9", "2", "2"]);
}

#[test]
fn test_list_get_or_else_default_and_null_lookup() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableListOf(1, 2, 3)
            println(values.getOrElse(1) { -1 })
            println(values.getOrNull(5) ?: -1)
            println(values.getOrElse(5) { -1 })
            println(values.getOrNull(0))
        }
    "#);
    assert_eq!(out, &["2", "-1", "-1", "1"]);
}

#[test]
fn test_list_sublist_mutates_parent_when_cleared() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableListOf("a", "b", "c", "d")
            val window = values.subList(1, 3)
            window.clear()
            println(values.joinToString(","))
            println(window.size)
        }
    "#);
    assert_eq!(out, &["a,d", "0"]);
}

#[test]
fn test_list_sublist_mutates_parent_when_mutated() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableListOf(1, 2, 3, 4, 5)
            val window = values.subList(1, 4)
            window[1] = 30
            window.removeAt(2)
            println(values.joinToString(","))
            println(window.size)
        }
    "#);
    assert_eq!(out, &["1,2,30,5", "2"]);
}

#[test]
fn test_list_sublist_invalid_range_is_runtime_error() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableListOf(1, 2, 3)
            try {
                values.subList(3, 1)
                println("no-error")
            } catch (e: Exception) {
                println("error")
            }
        }
    "#);
    assert_eq!(out, &["error"]);
}

#[test]
fn test_set_to_list_roundtrip_keeps_distinct_elements() {
    let out = run_prints(r#"
        fun main() {
            val source = mutableListOf(1, 2, 2, 3, 1)
            val unique = source.toSet()
            val back = unique.toMutableList()
            back.add(4)
            println(unique.size)
            println(back.size)
            println(back.contains(4))
            println(source.size)
        }
    "#);
    assert_eq!(out, &["3", "4", "true", "5"]);
}

#[test]
fn test_list_reversed_returns_independent_copy() {
    let out = run_prints(r#"
        fun main() {
            val original = mutableListOf(1, 2, 3)
            val reversed = original.reversed()
            println(reversed.joinToString(","))
            original[0] = 9
            println(reversed.joinToString(","))
            println(original.joinToString(","))
        }
    "#);
    assert_eq!(out, &["3,2,1", "3,2,1", "9,2,3"]);
}

#[test]
fn test_list_plus_creates_new_list_without_mutating_source() {
    let out = run_prints(r#"
        fun main() {
            val head = mutableListOf(1, 2)
            val merged = head + listOf(3, 4)
            println(merged.joinToString(","))
            println(head.size)
            println(merged.size)
        }
    "#);
    assert_eq!(out, &["1,2,3,4", "2", "4"]);
}

#[test]
fn test_map_get_value_throws_when_missing_key() {
    let out = run_prints(r#"
        fun main() {
            val scores = mapOf("a" to 1, "b" to 2)
            try {
                println(scores.getValue("z"))
            } catch (e: NoSuchElementException) {
                println("missing")
            }
        }
    "#);
    assert_eq!(out, &["missing"]);
}

#[test]
fn test_map_get_or_default_and_or_else() {
    let out = run_prints(r#"
        fun main() {
            val scores = mapOf("a" to 10, "b" to 20)
            println(scores.getOrDefault("a", 0))
            println(scores.getOrDefault("c", 30))
            println(scores.getOrElse("b") { 0 })
            println(scores.getOrElse("c") { 99 })
        }
    "#);
    assert_eq!(out, &["10", "30", "20", "99"]);
}

#[test]
fn test_map_to_sorted_map_and_lookup() {
    let out = run_prints(r#"
        fun main() {
            val metrics = mapOf(3 to "c", 1 to "a", 2 to "b")
            val sorted = metrics.toSortedMap()
            println(sorted.keys.first())
            println(sorted.keys.last())
            println(sorted[2])
        }
    "#);
    assert_eq!(out, &["1", "3", "b"]);
}

#[test]
fn test_map_plus_operator_keeps_last_value_for_duplicate_key() {
    let out = run_prints(r#"
        fun main() {
            val left = mapOf("a" to 1, "b" to 2)
            val right = mapOf("b" to 9, "c" to 3)
            val merged = left + right
            println(merged.size)
            println(merged["b"])
            println(merged["c"])
        }
    "#);
    assert_eq!(out, &["3", "9", "3"]);
}

#[test]
fn test_map_minus_operator_removes_keys() {
    let out = run_prints(r#"
        fun main() {
            val base = mapOf("a" to 1, "b" to 2, "c" to 3)
            val narrowed = base - "b"
            println(narrowed.size)
            println(narrowed.containsKey("b"))
            println(narrowed["c"])
        }
    "#);
    assert_eq!(out, &["2", "false", "3"]);
}

#[test]
fn test_map_filter_to_map_and_keys_set() {
    let out = run_prints(r#"
        fun main() {
            val scores = mapOf("a" to 1, "b" to 4, "c" to 2, "d" to 5)
            val high = scores.filterValues { it >= 4 }
            println(high.size)
            println(high.keys.joinToString(","))
            println(high.values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["2", "b,d", "4,5"]);
}

#[test]
fn test_map_map_keys_and_values_transform() {
    let out = run_prints(r#"
        fun main() {
            val input = mapOf("a" to 1, "b" to 2)
            val keys = input.mapKeys { it.key.uppercase() }
            val values = input.mapValues { it.value * 10 }
            println(keys["A"])
            println(values["b"])
            println(keys.size)
        }
    "#);
    assert_eq!(out, &["1", "20", "2"]);
}

#[test]
fn test_map_put_if_absent_mutation_behavior() {
    let out = run_prints(r#"
        fun main() {
            val counts = mutableMapOf("a" to 1)
            println(counts.put("a", 9))
            println(counts.putIfAbsent("a", 11))
            println(counts.putIfAbsent("b", 2))
            println(counts["a"])
            println(counts["b"])
        }
    "#);
    assert_eq!(out, &["1", "1", "null", "9", "2"]);
}

#[test]
fn test_map_entries_as_set_size_and_mutation() {
    let out = run_prints(r#"
        fun main() {
            val inventory = mutableMapOf("x" to 1, "y" to 2)
            val entryView = inventory.entries
            var hasXOne = false
            for (entry in entryView) {
                if (entry.key == "x" && entry.value == 1) {
                    hasXOne = true
                }
            }
            println(entryView.size)
            inventory["z"] = 3
            println(entryView.size)
            println(hasXOne)
        }
    "#);
    assert_eq!(out, &["2", "3", "true"]);
}

#[test]
fn test_map_of_linked_preserves_assignment_insertion_order() {
    let out = run_prints(r#"
        fun main() {
            val history = linkedMapOf("first" to 1, "second" to 2, "third" to 3)
            var keys = ""
            for ((key, _) in history) {
                keys += key
            }
            println(keys)
            history["second"] = 4
            var keysAfter = ""
            for ((key, _) in history) {
                keysAfter += key
            }
            println(keysAfter)
        }
    "#);
    assert_eq!(out, &["firstsecondthird", "firstsecondthird"]);
}

#[test]
fn test_map_to_set_of_pairs_view() {
    let out = run_prints(r#"
        fun main() {
            val data = mapOf("a" to 1, "b" to 2)
            val pairs = data.toSet()
            println(pairs.size)
            println(pairs.contains(Pair("a", 1)))
            println(pairs.contains(Pair("b", 3)))
        }
    "#);
    assert_eq!(out, &["2", "true", "false"]);
}

#[test]
fn test_map_get_or_put_without_recomputation_on_existing_key() {
    let out = run_prints(r#"
        fun main() {
            var computed = 0
            val counts = mutableMapOf("a" to 1)
            println(counts.getOrPut("a") { computed += 1; 9 })
            println(counts.getOrPut("b") { computed += 1; 2 })
            println(counts["a"])
            println(counts["b"])
            println(computed)
        }
    "#);
    assert_eq!(out, &["1", "2", "1", "2", "1"]);
}

#[test]
fn test_map_with_default_keeps_original_map_without_side_updates() {
    let out = run_prints(r#"
        fun main() {
            val base = mapOf("known" to 5).withDefault { 77 }
            println(base.getValue("known"))
            println(base.getValue("missing"))
            println(base.toMap().containsKey("missing"))
        }
    "#);
    assert_eq!(out, &["5", "77", "false"]);
}

#[test]
fn test_map_put_all_from_pairs_sequence() {
    let out = run_prints(r#"
        fun main() {
            val source = mutableMapOf("a" to 1)
            source.putAll(listOf("a" to 9, "b" to 2).asSequence())
            println(source["a"])
            println(source["b"])
            println(source.size)
        }
    "#);
    assert_eq!(out, &["9", "2", "2"]);
}

#[test]
fn test_map_keys_view_remove_affects_map() {
    let out = run_prints(r#"
        fun main() {
            val source = mutableMapOf("a" to 1, "b" to 2, "c" to 3)
            val keys = source.keys
            keys.remove("b")
            println(source.size)
            println(source.containsKey("b"))
            println(source.containsKey("c"))
        }
    "#);
    assert_eq!(out, &["2", "false", "true"]);
}

#[test]
fn test_map_values_view_remove_by_value_affects_source() {
    let out = run_prints(r#"
        fun main() {
            val source = mutableMapOf("a" to 1, "b" to 2, "c" to 2)
            val values = source.values
            println(values.remove(2))
            println(source.size)
            println(source["c"] ?: -1)
        }
    "#);
    assert_eq!(out, &["true", "2", "-1"]);
}

#[test]
fn test_map_entries_iterator_mutation_is_fail_fast_for_structure_changes() {
    let out = run_prints(r#"
        fun main() {
            val source = mutableMapOf("a" to 1, "b" to 2)
            val iter = source.entries.iterator()
            println(iter.hasNext())
            println(iter.next().key)
            source["c"] = 3
            try {
                iter.next()
                println("no_fail")
            } catch (e: ConcurrentModificationException) {
                println("fail_fast")
            }
        }
    "#);
    assert_eq!(out, &["true", "a", "fail_fast"]);
}

#[test]
fn test_map_plus_assign_adds_and_removes_keys() {
    let out = run_prints(r#"
        fun main() {
            val source = linkedMapOf("a" to 1)
            source += mapOf("b" to 2, "c" to 3)
            source += mapOf("c" to 4)
            println(source["b"] + source["c"])
            println(source.size)
        }
    "#);
    assert_eq!(out, &["6", "3"]);
}

#[test]
fn test_map_minus_assign_removes_specified_key() {
    let out = run_prints(r#"
        fun main() {
            val source = linkedMapOf("a" to 1, "b" to 2, "c" to 3)
            source -= "b"
            println(source.size)
            println(source.containsKey("b"))
            println(source["a"] + source["c"])
        }
    "#);
    assert_eq!(out, &["2", "false", "4"]);
}

#[test]
fn test_map_to_mutable_map_is_copy_not_reference() {
    let out = run_prints(r#"
        fun main() {
            val base = mapOf("a" to 1, "b" to 2)
            val copy = base.toMutableMap()
            copy["a"] = 9
            println(base["a"])
            println(copy["a"])
            println(copy.size)
            println(base.size)
        }
    "#);
    assert_eq!(out, &["1", "9", "2", "2"]);
}

#[test]
fn test_map_contains_value_and_any() {
    let out = run_prints(r#"
        fun main() {
            val counters = mapOf("read" to 1, "write" to 2, "exec" to 0)
            println(counters.containsValue(2))
            println(counters.containsValue(3))
            println(counters.any { it.value > 1 })
            println(counters.all { it.key.isNotEmpty() })
        }
    "#);
    assert_eq!(out, &["true", "false", "true", "true"]);
}

#[test]
fn test_map_filter_keys_subset() {
    let out = run_prints(r#"
        fun main() {
            val metrics = mapOf("alpha" to 1, "beta" to 2, "gamma" to 3)
            val short = metrics.filterKeys { it.length == 4 }
            println(short.size)
            println(short["beta"])
            println(short.containsKey("alpha"))
        }
    "#);
    assert_eq!(out, &["2", "2", "false"]);
}

#[test]
fn test_map_get_or_else_calls_supplier_when_missing_only() {
    let out = run_prints(r#"
        fun main() {
            val scores = mapOf("a" to 1)
            var asked = 0
            val miss = scores.getOrElse("b") {
                asked += 1
                99
            }
            val hit = scores.getOrElse("a") {
                asked += 1
                88
            }
            println(miss)
            println(hit)
            println(asked)
        }
    "#);
    assert_eq!(out, &["99", "1", "1"]);
}

#[test]
fn test_map_get_or_put_default_does_not_override_existing() {
    let out = run_prints(r#"
        fun main() {
            val counters = mutableMapOf("x" to 1)
            println(counters.getOrPut("x") { 9 })
            println(counters["x"])
            println(counters.getOrPut("y") { 4 })
            println(counters["y"])
            println(counters.size)
        }
    "#);
    assert_eq!(out, &["1", "1", "4", "4", "2"]);
}
