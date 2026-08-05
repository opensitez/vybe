use crate::helpers::run_prints;

#[test]
fn test_map_filter_keys_and_values() {
    let out = run_prints(
        r#"
        fun main() {
            val data = mapOf("apple" to 5, "kiwi" to 3, "pear" to 8)
            val shortKeys = data.filterKeys { it.length < 5 }
            val highValues = data.filterValues { it >= 6 }
            println(shortKeys.keys.joinToString(","))
            println(highValues.keys.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["kiwi,pear", "pear"]);
}

#[test]
fn test_map_map_keys_rename_projection() {
    let out = run_prints(
        r#"
        fun main() {
            val original = mapOf("a" to 1, "b" to 2)
            val renamed = original.mapKeys { it.key + "!" }
            println(renamed.keys.joinToString(","))
            println(renamed["a!"] + renamed["b!"])
        }
    "#,
    );
    assert_eq!(out, &["a!,b!", "3"]);
}

#[test]
fn test_map_map_values_transform() {
    let out = run_prints(
        r#"
        fun main() {
            val source = mapOf("a" to 1, "b" to 2)
            val doubled = source.mapValues { (_, value) -> value * 3 }
            println(doubled["a"])
            println(doubled["b"])
        }
    "#,
    );
    assert_eq!(out, &["3", "6"]);
}

#[test]
fn test_map_get_or_default_for_missing_key() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mapOf("one" to 1, "two" to 2)
            println(map.getOrDefault("one", -1))
            println(map.getOrDefault("three", -1))
        }
    "#,
    );
    assert_eq!(out, &["1", "-1"]);
}

#[test]
fn test_map_get_or_else_with_supplier() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mapOf("one" to 1)
            val existing = map.getOrElse("one") { 99 }
            val missing = map.getOrElse("two") { 99 }
            println(existing)
            println(missing)
        }
    "#,
    );
    assert_eq!(out, &["1", "99"]);
}

#[test]
fn test_map_get_value_throws_when_missing() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mapOf("a" to 1)
            try {
                println(map.getValue("b"))
            } catch (e: NoSuchElementException) {
                println("missing")
            }
        }
    "#,
    );
    assert_eq!(out, &["missing"]);
}

#[test]
fn test_map_or_empty_for_nullable_source() {
    let out = run_prints(
        r#"
        fun main() {
            val maybe: Map<String, Int>? = null
            val safe = maybe.orEmpty()
            println(safe.isEmpty())
            println(safe.size)
        }
    "#,
    );
    assert_eq!(out, &["true", "0"]);
}

#[test]
fn test_map_contains_key_and_value() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mapOf("a" to 1, "b" to 2, "c" to 2)
            println(map.containsKey("b"))
            println(map.containsValue(2))
            println(map.containsValue(3))
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "false"]);
}

#[test]
fn test_map_count_and_any_all_none() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mapOf("a" to 1, "b" to 2, "c" to 3)
            println(map.count { it.value > 1 })
            println(map.any { it.key == "b" })
            println(map.all { it.value > 0 })
            println(map.none { it.key == "z" })
        }
    "#,
    );
    assert_eq!(out, &["2", "true", "true", "true"]);
}

#[test]
fn test_map_get_or_put_existing_preserves_value() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mutableMapOf("a" to 1)
            val first = map.getOrPut("a") { 9 }
            val second = map.getOrPut("b") { 9 }
            println(first)
            println(second)
            println(map["a"])
            println(map["b"])
        }
    "#,
    );
    assert_eq!(out, &["1", "9", "1", "9"]);
}

#[test]
fn test_map_put_if_absent() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mutableMapOf("a" to 1)
            val existing = map.putIfAbsent("a", 9)
            val added = map.putIfAbsent("b", 2)
            println(existing)
            println(added)
            println(map["a"])
            println(map["b"])
        }
    "#,
    );
    assert_eq!(out, &["1", "null", "1", "2"]);
}

#[test]
fn test_map_remove_key_and_value_signature() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            println(map.remove("a", 1))
            println(map.remove("b", 1))
            println(map.size)
            println(map["a"])
            println(map["b"])
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "1", "null", "2"]);
}

#[test]
fn test_map_plus_operator_merges_override() {
    let out = run_prints(
        r#"
        fun main() {
            val a = mapOf("a" to 1, "b" to 2)
            val b = mapOf("b" to 9, "c" to 3)
            val merged = a + b
            println(merged["a"])
            println(merged["b"])
            println(merged["c"])
            println(merged.size)
        }
    "#,
    );
    assert_eq!(out, &["1", "9", "3", "3"]);
}

#[test]
fn test_map_minus_operator_removes_key() {
    let out = run_prints(
        r#"
        fun main() {
            val source = mapOf("a" to 1, "b" to 2, "c" to 3)
            val reduced = source - "b"
            println(reduced.size)
            println(reduced.containsKey("b"))
            println(reduced["a"] + reduced["c"])
        }
    "#,
    );
    assert_eq!(out, &["2", "false", "4"]);
}

#[test]
fn test_map_put_all_from_iterable() {
    let out = run_prints(
        r#"
        fun main() {
            val base = mutableMapOf("a" to 1)
            base.putAll(listOf("b" to 2, "c" to 3))
            println(base["a"] + base["b"] + base["c"])
            println(base.size)
        }
    "#,
    );
    assert_eq!(out, &["6", "3"]);
}

#[test]
fn test_map_to_mutable_map_is_independent() {
    let out = run_prints(
        r#"
        fun main() {
            val source = mapOf("a" to 1, "b" to 2)
            val copy = source.toMutableMap()
            copy["a"] = 9
            println(source["a"])
            println(copy["a"])
            println(copy.size)
            println(source.size)
        }
    "#,
    );
    assert_eq!(out, &["1", "9", "2", "2"]);
}

#[test]
fn test_map_to_map_snapshot_is_not_reactive() {
    let out = run_prints(
        r#"
        fun main() {
            val source = mutableMapOf("a" to 1)
            val snap = source.toMap()
            source["a"] = 4
            source["b"] = 2
            println(snap["a"])
            println(snap.size)
            println(source.size)
        }
    "#,
    );
    assert_eq!(out, &["1", "1", "2"]);
}

#[test]
fn test_map_keys_and_values_views() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mapOf("a" to 1, "b" to 2, "c" to 3)
            println(map.keys.joinToString(","))
            println(map.values.sum())
            println(map.values.maxOrNull() ?: 0)
        }
    "#,
    );
    assert_eq!(out, &["a,b,c", "6", "3"]);
}

#[test]
fn test_map_entries_find_and_first() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mapOf("x" to 10, "y" to 11, "z" to 12)
            val found = map.entries.find { it.value == 11 }
            println(found?.key ?: "none")
            val first = map.entries.first()
            println(first.key + ":" + first.value)
        }
    "#,
    );
    assert_eq!(out, &["y", "x:10"]);
}

#[test]
fn test_map_entries_iteration_order_in_linked_map() {
    let out = run_prints(
        r#"
        fun main() {
            val map = linkedMapOf("first" to 1, "second" to 2, "third" to 3)
            var keys = ""
            var sum = 0
            for (entry in map.entries) {
                keys += entry.key
                sum += entry.value
            }
            println(keys)
            println(sum)
        }
    "#,
    );
    assert_eq!(out, &["firstsecondthird", "6"]);
}

#[test]
fn test_map_merge_duplicate_keys_keeps_last() {
    let out = run_prints(
        r#"
        fun main() {
            val pairs = listOf("a" to 1, "b" to 2, "a" to 3)
            val map = pairs.toMap()
            println(map["a"])
            println(map.size)
            println(map["b"])
        }
    "#,
    );
    assert_eq!(out, &["3", "2", "2"]);
}

#[test]
fn test_map_build_from_iterable_and_to_list() {
    let out = run_prints(
        r#"
        fun main() {
            val map = listOf("a" to 1, "b" to 2).toMap()
            val items = map.toList()
            println(items[0].first)
            println(items[1].second)
            println(items.size)
        }
    "#,
    );
    assert_eq!(out, &["a", "2", "2"]);
}

#[test]
fn test_map_associate_by_projection() {
    let out = run_prints(
        r#"
        fun main() {
            val words = listOf("one", "two", "three")
            val map = words.associateBy { it.first() }
            println(map['o'])
            println(map['t'])
        }
    "#,
    );
    assert_eq!(out, &["one", "three"]);
}

#[test]
fn test_map_grouping_by_value_parity() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mapOf("a" to 1, "b" to 2, "c" to 3, "d" to 4)
            val parity = map.entries.groupBy { it.value % 2 == 0 }
            println(parity[true]?.size ?: 0)
            println(parity[false]?.size ?: 0)
        }
    "#,
    );
    assert_eq!(out, &["2", "2"]);
}

#[test]
fn test_map_update_and_remove_cycle() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mutableMapOf("x" to 1)
            map["x"] = 2
            map.remove("x")
            map["x"] = 5
            println(map["x"])
            println(map.size)
        }
    "#,
    );
    assert_eq!(out, &["5", "1"]);
}

#[test]
fn test_map_for_each_accumulation() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mapOf("a" to 1, "b" to 2, "c" to 3)
            var sumKeys = 0
            var sumValues = 0
            map.forEach { _, value ->
                sumValues += value
            }
            for (_ in map) {
                sumKeys += 1
            }
            println(sumValues)
            println(sumKeys)
        }
    "#,
    );
    assert_eq!(out, &["6", "3"]);
}

#[test]
fn test_map_empty_defaults_and_is_empty() {
    let out = run_prints(
        r#"
        fun main() {
            val map = emptyMap<String, Int>()
            println(map.isEmpty())
            println(map.getOrDefault("x", 5))
            println(map.orEmpty().size)
        }
    "#,
    );
    assert_eq!(out, &["true", "5", "0"]);
}

#[test]
fn test_map_with_nullable_keys_and_values() {
    let out = run_prints(
        r#"
        fun main() {
            val map: Map<String?, Int?> = mapOf(null to 5, "x" to null)
            println(map[null])
            println(map.containsKey(null))
            println(map["x"] ?: -1)
            println(map.getOrElse(null) { -1 })
        }
    "#,
    );
    assert_eq!(out, &["5", "true", "-1", "5"]);
}

#[test]
fn test_map_nullable_key_lookup_missing() {
    let out = run_prints(
        r#"
        fun main() {
            val map: Map<Int?, String> = mapOf(null to "nil", 1 to "one")
            println(map[2] ?: "none")
            println(map[null])
            println(map.containsKey(2))
        }
    "#,
    );
    assert_eq!(out, &["none", "nil", "false"]);
}

#[test]
fn test_map_data_class_key_equality() {
    let out = run_prints(
        r#"
        data class Key(val id: Int)

        fun main() {
            val map = mapOf(Key(1) to "first", Key(2) to "second")
            println(map[Key(1)])
            println(map[Key(2)])
        }
    "#,
    );
    assert_eq!(out, &["first", "second"]);
}

#[test]
fn test_map_sorted_copy_orders_keys() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mapOf("c" to 3, "a" to 1, "b" to 2)
            val sorted = map.toSortedMap()
            println(sorted.keys.joinToString(","))
            println(sorted.values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["a,b,c", "1,2,3"]);
}

#[test]
fn test_map_get_or_put_complex_default() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mutableMapOf("seed" to mutableListOf(1))
            map.getOrPut("seed") { mutableListOf() }.add(2)
            val values = map.getOrPut("fresh") { mutableListOf(9) }
            values.add(10)
            println(map["seed"]?.size)
            println(map["seed"]?.get(1))
            println(map["fresh"]?.size)
        }
    "#,
    );
    assert_eq!(out, &["2", "2", "2"]);
}

#[test]
fn test_map_clear_and_repopulate() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            map.clear()
            println(map.isEmpty())
            map["z"] = 10
            println(map.size)
            println(map["z"])
        }
    "#,
    );
    assert_eq!(out, &["true", "1", "10"]);
}

#[test]
fn test_map_put_returns_previous_value() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            println(map.put("a", 5))
            println(map.put("c", 3))
            println(map["a"] + map["c"])
        }
    "#,
    );
    assert_eq!(out, &["1", "null", "8"]);
}

#[test]
fn test_map_remove_returns_value_or_null() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            println(map.remove("a"))
            println(map.remove("missing"))
            println(map.size)
        }
    "#,
    );
    assert_eq!(out, &["1", "null", "1"]);
}

#[test]
fn test_map_replace_overwrites_only_when_present() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            println(map.replace("a", 7))
            println(map.replace("c", 9))
            println(map["a"])
            println(map["c"] ?: -1)
        }
    "#,
    );
    assert_eq!(out, &["1", "null", "7", "-1"]);
}

#[test]
fn test_map_conditional_replace_with_match_value() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            println(map.replace("a", 1, 9))
            println(map.replace("a", 2, 10))
            println(map["a"])
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "9"]);
}

#[test]
fn test_map_put_all_overwrites_existing_entries() {
    let out = run_prints(
        r#"
        fun main() {
            val base = mutableMapOf("a" to 1, "b" to 2)
            base.putAll(mapOf("b" to 4, "c" to 5))
            println(base["b"])
            println(base["c"])
            println(base.size)
        }
    "#,
    );
    assert_eq!(out, &["4", "5", "3"]);
}

#[test]
fn test_map_put_mutates_size_and_includes_negative_key() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mutableMapOf<Int, Int>()
            map.put(-1, 10)
            map[-2] = 20
            println(map.size)
            println(map[-1] + map[-2])
        }
    "#,
    );
    assert_eq!(out, &["2", "30"]);
}

#[test]
fn test_map_keys_and_values_views_reflect_mutations() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            val keys = map.keys
            val values = map.values
            map["a"] = 4
            map["c"] = 3
            println(keys.contains("c"))
            println(values.any { it == 4 })
            println(values.sum())
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "9"]);
}

#[test]
fn test_map_mutable_minus_operator_creates_snapshot() {
    let out = run_prints(
        r#"
        fun main() {
            val source = mutableMapOf("a" to 1, "b" to 2, "c" to 3)
            val reduced = source - "b"
            source["d"] = 4
            println(reduced.containsKey("d"))
            println(source["d"])
            println(reduced.size)
        }
    "#,
    );
    assert_eq!(out, &["false", "4", "2"]);
}

#[test]
fn test_map_filter_not_removes_matching_entries() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mapOf("a" to 1, "b" to 2, "c" to 3)
            val evenValues = map.filterNot { entry -> entry.value % 2 == 0 }
            println(evenValues.keys.joinToString(","))
            println(evenValues.size)
        }
    "#,
    );
    assert_eq!(out, &["a,c", "2"]);
}

#[test]
fn test_map_keys_view_removal_mutates_backing_map() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2, "c" to 3)
            val keys = map.keys
            keys.remove("b")
            println(map.size)
            println(map.containsKey("b"))
            println(map["a"] + map["c"])
        }
    "#,
    );
    assert_eq!(out, &["2", "false", "4"]);
}

#[test]
fn test_map_values_view_removal_mutates_backing_map() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2, "c" to 2)
            val values = map.values
            println(values.remove(2))
            println(map.size)
            println(map.values.sum())
        }
    "#,
    );
    assert_eq!(out, &["true", "2", "3"]);
}

#[test]
fn test_map_iterator_is_fail_fast_when_mutated() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            val iter = map.iterator()
            println(iter.hasNext())
            println(iter.next().key)
            map["c"] = 3
            try {
                iter.next()
                println("no_fail")
            } catch (e: ConcurrentModificationException) {
                println("fail_fast")
            }
            println(map.size)
        }
    "#,
    );
    assert_eq!(out, &["true", "a", "fail_fast", "3"]);
}

#[test]
fn test_map_with_default_does_not_mutate_source() {
    let out = run_prints(
        r#"
        fun main() {
            val source = mapOf("one" to 1).withDefault { 99 }
            println(source.getValue("one"))
            println(source.getValue("two"))
            println(source.toMap().containsKey("two"))
        }
    "#,
    );
    assert_eq!(out, &["1", "99", "false"]);
}

#[test]
fn test_map_plus_assign_mutates_same_map() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mutableMapOf("a" to 1)
            map += mapOf("b" to 2, "c" to 3)
            println(map.size)
            println(map["b"] + map["c"])
        }
    "#,
    );
    assert_eq!(out, &["3", "5"]);
}

#[test]
fn test_map_minus_assign_keeps_non_removed_entries() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2, "c" to 3)
            map -= "b"
            println(map.size)
            println(map.containsKey("b"))
            println(map["a"] + map["c"])
        }
    "#,
    );
    assert_eq!(out, &["2", "false", "4"]);
}

#[test]
fn test_map_sorted_with_custom_comparator() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mapOf("a" to 1, "b" to 2, "c" to 3)
            val reversed = map.toSortedMap(compareByDescending { it })
            println(reversed.keys.joinToString(","))
            println(reversed.values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["c,b,a", "3,2,1"]);
}

#[test]
fn test_map_get_or_put_without_recomputing_for_present_key() {
    let out = run_prints(
        r#"
        fun main() {
            var computed = 0
            val map = mutableMapOf("present" to 1)
            val value1 = map.getOrPut("present") {
                computed += 1
                99
            }
            val value2 = map.getOrPut("missing") {
                computed += 1
                77
            }
            println(value1)
            println(value2)
            println(computed)
        }
    "#,
    );
    assert_eq!(out, &["1", "77", "1"]);
}

#[test]
fn test_map_associate_by_to_populates_destination() {
    let out = run_prints(
        r#"
        fun main() {
            val words = listOf("k", "ka", "kb")
            val byLength = mutableMapOf<Int, String>()
            words.associateByTo(byLength, { it.length })
            println(byLength[1])
            println(byLength[2])
        }
    "#,
    );
    assert_eq!(out, &["k", "kb"]);
}

#[test]
fn test_map_group_by_to_destination_merges_duplicates() {
    let out = run_prints(
        r#"
        fun main() {
            val words = listOf("a", "bb", "cc", "ddd", "eee")
            val groups = mutableMapOf<Int, MutableList<String>>()
            words.groupByTo(groups, { it.length })
            println(groups[1]?.size)
            println(groups[2]?.size)
            println(groups[3]?.size)
        }
    "#,
    );
    assert_eq!(out, &["1", "2", "2"]);
}

#[test]
fn test_map_associate_with_duplicate_keys_keeps_last_entry() {
    let out = run_prints(
        r#"
        fun main() {
            val source = listOf("a" to 1, "b" to 2, "a" to 3, "b" to 4)
            val map = source.associate { it.first to it.second }
            println(map["a"])
            println(map["b"])
            println(map.size)
        }
    "#,
    );
    assert_eq!(out, &["3", "4", "2"]);
}

#[test]
fn test_map_to_mutable_map_is_snapshot_insertion_independent() {
    let out = run_prints(
        r#"
        fun main() {
            val base = mapOf("x" to 1, "y" to 2)
            val copied = base.toMutableMap()
            copied["z"] = 3
            println(base["z"])
            println(copied["z"])
            println(copied.size)
            println(base.size)
        }
    "#,
    );
    assert_eq!(out, &["null", "3", "3", "2"]);
}

#[test]
fn test_map_iterator_remove_during_iteration() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2, "c" to 3, "d" to 4)
            val iter = map.entries.iterator()
            while (iter.hasNext()) {
                val current = iter.next()
                if (current.value % 2 == 0) {
                    iter.remove()
                }
            }
            println(map.size)
            println(map["b"])
            println(map["d"])
            println(map["a"] + map["c"])
        }
    "#,
    );
    assert_eq!(out, &["2", "null", "null", "4"]);
}

#[test]
fn test_map_filter_values_retains_only_matching_and_preserves_order() {
    let out = run_prints(
        r#"
        fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2, "c" to 3)
            val filtered = map.filterValues { it >= 2 }
            println(filtered.keys.joinToString(","))
            println(filtered.values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["b,c", "2,3"]);
}

#[test]
fn test_map_map_values_then_filter_keys_chain() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mapOf("a" to 1, "bb" to 2, "ccc" to 3)
            val result = map
                .mapValues { it.value * 2 }
                .filterKeys { it.length <= 2 }
            println(result["a"])
            println(result["bb"])
            println(result["ccc"] ?: -1)
        }
    "#,
    );
    assert_eq!(out, &["2", "4", "-1"]);
}

#[test]
fn test_map_count_and_sum_from_entries_projection() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mapOf("a" to 1, "b" to 2, "c" to 3)
            val sum = map.entries.map { it.value }.sum()
            val count = map.entries.filter { it.key != "b" }.count()
            println(sum)
            println(count)
        }
    "#,
    );
    assert_eq!(out, &["6", "2"]);
}

#[test]
fn test_map_join_to_string_with_entry_shape() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mapOf("a" to 1, "b" to 2)
            val formatted = map.entries.joinToString("|") { "${it.key}:${it.value}" }
            println(formatted)
        }
    "#,
    );
    assert_eq!(out, &["a:1|b:2"]);
}

#[test]
fn test_map_merge_without_overwrite_keeps_existing() {
    let out = run_prints(
        r#"
        fun main() {
            val base = linkedMapOf("a" to 1, "b" to 2)
            val extras = mapOf("b" to 20, "c" to 3)
            val merged = extras + base
            println(merged["a"])
            println(merged["b"])
            println(merged["c"])
        }
    "#,
    );
    assert_eq!(out, &["1", "2", "3"]);
}

#[test]
fn test_map_plus_assign_and_minus_assign_stability() {
    let out = run_prints(
        r#"
        fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2, "c" to 3)
            map += mapOf("d" to 4)
            map -= "b"
            println(map.size)
            println(map["a"] + (map["c"] ?: 0) + (map["d"] ?: 0))
            println(map["b"] ?: -1)
        }
    "#,
    );
    assert_eq!(out, &["3", "8", "-1"]);
}
