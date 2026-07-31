use crate::helpers::run_prints;

#[test]
fn test_linked_list_preserves_insertion_order() {
    let out = run_prints(r#"
        fun main() {
            val list = mutableListOf(3, 1, 2)
            println(list.joinToString(","))
            list.add(0, 9)
            println(list.joinToString(","))
        }
    "#);
    assert_eq!(out, &["3,1,2", "9,3,1,2"]);
}

#[test]
fn test_linked_set_preserves_insertion_order() {
    let out = run_prints(r#"
        fun main() {
            val set = LinkedHashSet<Int>()
            set.add(3)
            set.add(1)
            set.add(2)
            println(set.joinToString(","))
            set.remove(1)
            set.add(1)
            println(set.joinToString(","))
        }
    "#);
    assert_eq!(out, &["3,1,2", "3,2,1"]);
}

#[test]
fn test_linked_map_insertion_order_for_keys() {
    let out = run_prints(r#"
        fun main() {
            val map = LinkedHashMap<String, Int>()
            map["b"] = 2
            map["a"] = 1
            map["c"] = 3
            println(map.keys.joinToString(","))
        }
    "#);
    assert_eq!(out, &["b,a,c"]);
}

#[test]
fn test_linked_map_reinsertion_moves_key_to_last() {
    let out = run_prints(r#"
        fun main() {
            val map = LinkedHashMap<String, Int>()
            map["a"] = 1
            map["b"] = 2
            map["a"] = 9
            println(map.keys.joinToString(","))
            println(map["a"])
        }
    "#);
    assert_eq!(out, &["a,b", "9"]);
}

#[test]
fn test_map_iteration_over_entries_order() {
    let out = run_prints(r#"
        fun main() {
            val map = linkedMapOf("first" to 1, "second" to 2, "third" to 3)
            println(map.entries.joinToString(";") { it.key })
            println(map.entries.joinToString(";") { it.value.toString() })
        }
    "#);
    assert_eq!(out, &["first,second,third", "1,2,3"]);
}

#[test]
fn test_list_set_intersection_preserves_list_order_of_first() {
    let out = run_prints(r#"
        fun main() {
            val a = listOf(1, 2, 3, 4)
            val b = setOf(4, 2)
            val out = a.filter { b.contains(it) }
            println(out.joinToString(","))
        }
    "#);
    assert_eq!(out, &["2,4"]);
}

#[test]
fn test_set_to_list_keeps_iteration_order() {
    let out = run_prints(r#"
        fun main() {
            val set = linkedSetOf(4, 1, 3)
            println(set.toList().joinToString(","))
            println(set.toMutableList().sorted().joinToString(","))
        }
    "#);
    assert_eq!(out, &["4,1,3", "1,3,4"]);
}

#[test]
fn test_sorted_map_orders_by_key() {
    let out = run_prints(r#"
        fun main() {
            val map = java.util.TreeMap<String, Int>()
            map["b"] = 2
            map["a"] = 1
            map["c"] = 3
            println(map.keys.joinToString(","))
            println(map.values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["a,b,c", "1,2,3"]);
}

#[test]
fn test_sorted_set_orders() {
    let out = run_prints(r#"
        fun main() {
            val set = java.util.TreeSet<Int>()
            set.add(3)
            set.add(1)
            set.add(2)
            println(set.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,2,3"]);
}

#[test]
fn test_sorted_set_descending_order() {
    let out = run_prints(r#"
        fun main() {
            val set = java.util.TreeSet<Int>()
            set.add(1); set.add(3); set.add(2)
            val values = set.descendingSet()
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["3,2,1"]);
}

#[test]
fn test_navigable_map_next_previous_entry() {
    let out = run_prints(r#"
        fun main() {
            val map = java.util.TreeMap<Int, String>()
            map[1] = "one"
            map[3] = "three"
            map[5] = "five"
            println(map.higherKey(3))
            println(map.lowerKey(3))
        }
    "#);
    assert_eq!(out, &["5", "1"]);
}

#[test]
fn test_sequence_of_list_indices() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf("a", "b", "c")
            val indices = values.indices.toList()
            println(indices.joinToString(","))
            println(values[indices.first()])
            println(values[indices.last()])
        }
    "#);
    assert_eq!(out, &["0,1,2", "a", "c"]);
}

#[test]
fn test_map_keys_view_is_ordered_for_linked() {
    let out = run_prints(r#"
        fun main() {
            val map = linkedMapOf(1 to "a", 2 to "b", 3 to "c")
            val view = map.keys
            println(view.joinToString(","))
            map.remove(2)
            map[4] = "d"
            println(view.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,2,3", "1,3,4"]);
}

#[test]
fn test_map_values_view_reflects_updates() {
    let out = run_prints(r#"
        fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            val values = map.values
            map["a"] = 9
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["9,2"]);
}

#[test]
fn test_list_flat_map_preserves_order() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(listOf(1, 2), listOf(3), listOf(4, 5))
            println(values.flatMap { it }.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,2,3,4,5"]);
}

#[test]
fn test_flatten_map_order() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(listOf(2, 1), listOf(4, 3))
            val out = values.flatten()
            println(out.joinToString(","))
        }
    "#);
    assert_eq!(out, &["2,1,4,3"]);
}

#[test]
fn test_map_entries_iteration_order_after_clear() {
    let out = run_prints(r#"
        fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            map.clear()
            map["c"] = 3
            map["d"] = 4
            println(map.entries.joinToString(",") { it.key })
        }
    "#);
    assert_eq!(out, &["c,d"]);
}

#[test]
fn test_list_retain_and_drop_order() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(1, 2, 3, 4, 5)
            println(values.filter { it % 2 == 1 }.joinToString(","))
            println(values.dropWhile { it < 4 }.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,3,5", "4,5"]);
}

#[test]
fn test_set_retain_all_order_from_list() {
    let out = run_prints(r#"
        fun main() {
            val set = linkedSetOf(1, 2, 3, 4)
            set.retainAll(listOf(4, 2))
            println(set.joinToString(","))
        }
    "#);
    assert_eq!(out, &["2,4"]);
}

#[test]
fn test_map_put_all_preserves_new_order_for_new_keys() {
    let out = run_prints(r#"
        fun main() {
            val map = linkedMapOf("a" to 1)
            map.putAll(mapOf("c" to 3, "b" to 2))
            println(map.keys.joinToString(","))
        }
    "#);
    assert_eq!(out, &["a,c,b"]);
}

#[test]
fn test_list_sorted_by_comparator_stable() {
    let out = run_prints(r#"
        data class Pair(val left: Int, val right: Int)

        fun main() {
            val list = listOf(Pair(1, 2), Pair(1, 1), Pair(0, 3))
            val sorted = list.sortedWith(compareBy<Pair> { it.left }.thenBy { it.right })
            println(sorted.map { "${'$'}{it.left}:${'$'}{it.right}" }.joinToString("|"))
        }
    "#);
    assert_eq!(out, &["0:3|1:1|1:2"]);
}

#[test]
fn test_set_to_string_preserves_deterministic_for_linked_set() {
    let out = run_prints(r#"
        fun main() {
            val set = linkedSetOf("x", "y", "z")
            println(set.toString())
        }
    "#);
    assert_eq!(out, &["[x, y, z]"]);
}

#[test]
fn test_list_partition_preserves_relative_order() {
    let out = run_prints(r#"
        fun main() {
            val list = listOf(1, 2, 3, 4, 5)
            val (a, b) = list.partition { it % 2 == 0 }
            println(a.joinToString(","))
            println(b.joinToString(","))
        }
    "#);
    assert_eq!(out, &["2,4", "1,3,5"]);
}

#[test]
fn test_collection_concatenation_order() {
    let out = run_prints(r#"
        fun main() {
            val a = listOf(1, 2)
            val b = listOf(3, 4)
            println((a + b).joinToString(","))
            println(a.plus(b).joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,2,3,4", "1,2,3,4"]);
}

#[test]
fn test_map_keys_to_list_after_mutation() {
    let out = run_prints(r#"
        fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            val keys = map.keys.toMutableList()
            map["c"] = 3
            println(keys.joinToString(","))
            println(map.keys.joinToString(","))
        }
    "#);
    assert_eq!(out, &["a,b", "a,b,c"]);
}

#[test]
fn test_map_to_list_round_trip_order() {
    let out = run_prints(r#"
        fun main() {
            val list = listOf("a" to 1, "b" to 2, "c" to 3)
            val map = linkedMapOf<String, Int>()
            map.putAll(list.toMap())
            val rebuilt = map.toList()
            println(rebuilt.joinToString("|") { "${'$'}{it.first}:${'$'}{it.second}" })
        }
    "#);
    assert_eq!(out, &["a:1|b:2|c:3"]);
}
