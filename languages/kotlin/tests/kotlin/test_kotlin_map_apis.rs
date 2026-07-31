kotlin_run_cases! {
    test_map_basic_lookup => (r##"
        fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2, "c" to 3)
            println(map["a"])
            println(map["c"])
            println(map.size)
        }
    "##, &[
        "1",
        "3",
        "3",
    ]),
    test_map_duplicate_key_last_wins => (r##"
        fun main() {
            val map = mapOf("x" to 1, "x" to 9, "y" to 4)
            println(map["x"])
            println(map.size)
        }
    "##, &[
        "9",
        "2",
    ]),
    test_map_contains_key_value => (r##"
        fun main() {
            val map = mapOf("a" to 1, "b" to 2)
            println(map.containsKey("a"))
            println(map.containsKey("z"))
            println(map.containsValue(2))
        }
    "##, &[
        "true",
        "false",
        "true",
    ]),
    test_map_get_or_default => (r##"
        fun main() {
            val map = mapOf("a" to 1, "b" to 2)
            println(map.getOrDefault("a", 9))
            println(map.getOrDefault("z", 9))
        }
    "##, &[
        "1",
        "9",
    ]),
    test_map_get_or_else => (r##"
        fun main() {
            val map = mapOf("a" to 1)
            println(map.getOrElse("a") { 0 })
            println(map.getOrElse("z") { 7 })
        }
    "##, &[
        "1",
        "7",
    ]),
    test_map_empty_and_size => (r##"
        fun main() {
            val map = emptyMap<String, Int>()
            println(map.isEmpty())
            println(map.isNotEmpty())
            println(map.size)
        }
    "##, &[
        "true",
        "false",
        "0",
    ]),
    test_map_keys_view_order => (r##"
        fun main() {
            val map = linkedMapOf("b" to 2, "a" to 1, "c" to 3)
            println(map.keys.joinToString(","))
            println(map.values.joinToString(","))
        }
    "##, &[
        "b,a,c",
        "2,1,3",
    ]),
    test_map_mutable_update => (r##"
        fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            map["a"] = 8
            map["c"] = 4
            println(map["a"])
            println(map["c"])
            println(map.size)
        }
    "##, &[
        "8",
        "4",
        "3",
    ]),
    test_map_remove_key => (r##"
        fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2, "c" to 3)
            val removed = map.remove("b")
            println(removed)
            println(map.size)
            println(map.containsKey("b"))
        }
    "##, &[
        "2",
        "2",
        "false",
    ]),
    test_map_remove_missing => (r##"
        fun main() {
            val map = mutableMapOf("a" to 1)
            val removed = map.remove("z")
            println(removed == null)
            println(map.size)
        }
    "##, &[
        "true",
        "1",
    ]),
    test_map_plus_merges_overwrites => (r##"
        fun main() {
            val a = linkedMapOf("a" to 1, "b" to 2)
            val b = linkedMapOf("b" to 9, "c" to 3)
            val merged = a + b
            println(merged["a"])
            println(merged["b"])
            println(merged["c"])
        }
    "##, &[
        "1",
        "9",
        "3",
    ]),
    test_map_minus_removes_keys => (r##"
        fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2, "c" to 3)
            val reduced = map - "b"
            println(reduced.size)
            println(reduced.containsKey("b"))
        }
    "##, &[
        "2",
        "false",
    ]),
    test_map_filter_keys => (r##"
        fun main() {
            val map = linkedMapOf("aa" to 1, "bb" to 2, "abc" to 3)
            val filtered = map.filterKeys { it.length > 2 }
            println(filtered.size)
            println(filtered["abc"])
        }
    "##, &[
        "1",
        "3",
    ]),
    test_map_filter_values => (r##"
        fun main() {
            val map = linkedMapOf("a" to 1, "b" to 4, "c" to 2)
            val filtered = map.filterValues { it >= 3 }
            println(filtered.size)
            println(filtered["b"])
            println(filtered.containsKey("a"))
        }
    "##, &[
        "1",
        "4",
        "false",
    ]),
    test_map_map_keys => (r##"
        fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            val transformed = map.mapKeys { it.key + "#" }
            println(transformed.keys.joinToString(","))
            println(transformed.size)
        }
    "##, &[
        "a#,b#",
        "2",
    ]),
    test_map_map_values => (r##"
        fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            val transformed = map.mapValues { it.value * 10 }
            println(transformed["a"])
            println(transformed["b"])
        }
    "##, &[
        "10",
        "20",
    ]),
    test_map_entries_iteration => (r##"
        fun main() {
            val map = linkedMapOf("x" to 1, "y" to 2)
            var seen = ""
            for (entry in map.entries) {
                seen += entry.key + entry.value
            }
            println(seen)
        }
    "##, &[
        "x1y2",
    ]),
    test_map_for_each_collect => (r##"
        fun main() {
            val map = linkedMapOf("a" to 3, "b" to 5)
            var total = 0
            map.forEach { _, value -> total += value }
            println(total)
        }
    "##, &[
        "8",
    ]),
    test_map_copy_to_mutable_keeps_values => (r##"
        fun main() {
            val source = linkedMapOf("a" to 1, "b" to 2)
            val copy = source.toMutableMap()
            copy.put("c", 3)
            println(source.size)
            println(copy.size)
            println(copy["c"])
        }
    "##, &[
        "2",
        "3",
        "3",
    ]),
    test_map_to_list_of_pairs => (r##"
        fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            val pairs = map.toList()
            println(pairs.size)
            println(pairs[0].first)
            println(pairs[1].second)
        }
    "##, &[
        "2",
        "a",
        "2",
    ]),
    test_map_count_values_gt_one => (r##"
        fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2, "c" to 3)
            val overOne = map.filterValues { it > 1 }.size
            val sum = map.values.sum()
            println(overOne)
            println(sum)
        }
    "##, &[
        "2",
        "6",
    ]),
    test_map_contains_all_of_entries => (r##"
        fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2, "c" to 3)
            val hasA = map.containsKey("a")
            val hasB = map.containsKey("b")
            println(hasA && hasB)
        }
    "##, &[
        "true",
    ]),
    test_map_to_string_has_entries => (r##"
        fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            println(map.toString())
        }
    "##, &[
        "{a=1, b=2}",
    ]),
    test_map_keys_any_match => (r##"
        fun main() {
            val map = linkedMapOf("alpha" to 1, "beta" to 2, "gamma" to 3)
            val hasLong = map.keys.any { it.length > 4 }
            val hasShort = map.keys.none { it.length > 6 }
            println(hasLong)
            println(hasShort)
        }
    "##, &[
        "true",
        "true",
    ]),
    test_map_values_projection_sum => (r##"
        fun main() {
            val map = linkedMapOf("a" to 10, "b" to 20, "c" to 30)
            val values = map.values.toMutableList()
            values[1] = 25
            println(values.sum())
        }
    "##, &[
        "65",
    ]),
    test_map_retain_order_after_update => (r##"
        fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2, "c" to 3)
            map["b"] = 9
            println(map.keys.joinToString(","))
            println(map["b"])
        }
    "##, &[
        "a,b,c",
        "9",
    ]),
    test_map_replace_value => (r##"
        fun main() {
            val map = linkedMapOf("x" to 1)
            val previous = map.put("x", 7)
            println(previous)
            println(map["x"])
        }
    "##, &[
        "1",
        "7",
    ]),
    test_map_remove_and_replace_chain => (r##"
        fun main() {
            val map = linkedMapOf("x" to 1, "y" to 2)
            map.remove("x")
            map["z"] = 3
            println(map.size)
            println(map.containsKey("x"))
            println(map["z"])
        }
    "##, &[
        "2",
        "false",
        "3",
    ]),
    test_map_with_pair_array_to_map => (r##"
        fun main() {
            val pairs = arrayOf(Pair("n", 1), Pair("m", 2), Pair("n", 4))
            val map = pairs.toMap()
            println(map.size)
            println(map["n"])
        }
    "##, &[
        "2",
        "4",
    ]),
    test_map_is_not_empty_after_clear => (r##"
        fun main() {
            val map = linkedMapOf("a" to 1)
            val before = map.isNotEmpty()
            map.clear()
            println(before)
            println(map.isEmpty())
        }
    "##, &[
        "true",
        "true",
    ]),
    test_map_keys_to_set_size => (r##"
        fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            println(map.keys.toSet().size)
            println(map.keys.toSet().toList().joinToString(","))
        }
    "##, &[
        "2",
        "a,b",
    ]),
    test_map_or_empty_is_empty_map => (r##"
        fun main() {
            val map: Map<String, Int> = mapOf()
            val empty = map.orEmpty()
            println(empty.isEmpty())
            println(empty.size)
        }
    "##, &[
        "true",
        "0",
    ]),
    test_map_sum_values_of_mutable_map => (r##"
        fun main() {
            val map = mutableMapOf("a" to 2, "b" to 4)
            println((map["a"] ?: 0) + (map["b"] ?: 0))
            map["a"] = map["a"]!! + 3
            println(map["a"])
        }
    "##, &[
        "6",
        "5",
    ]),
    test_map_map_entry_copy_roundtrip => (r##"
        fun main() {
            val map = linkedMapOf("x" to 9)
            val round = linkedMapOf(map.entries.first().toPair())
            println(round["x"])
            println(round.size)
        }
    "##, &[
        "9",
        "1",
    ]),
}
