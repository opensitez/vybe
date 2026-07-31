use crate::helpers::run_prints;

#[test]
fn test_build_list_collects_all_items() {
    let out = run_prints(r#"
        fun main() {
            val values = buildList {
                add(1)
                add(2)
                add(3)
            }
            println(values.size)
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["3", "1,2,3"]);
}

#[test]
fn test_build_list_index_access_and_mutability_of_read_only_view() {
    let out = run_prints(r#"
        fun main() {
            val values = buildList(4) {
                for (i in 0 until 4) {
                    add(i * 3)
                }
            }
            println(values[0])
            println(values[2])
            println(values.size)
        }
    "#);
    assert_eq!(out, &["0", "6", "4"]);
}

#[test]
fn test_build_list_can_skip_by_condition() {
    let out = run_prints(r#"
        fun main() {
            val values = buildList {
                for (v in 1..6) {
                    if (v % 2 == 0) add(v)
                }
            }
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["2,4,6"]);
}

#[test]
fn test_build_list_size_hint_affects_internal_growth() {
    let out = run_prints(r#"
        fun main() {
            val values = buildList(2) {
                add("a")
                add("b")
                add("c")
            }
            println(values.size)
            println(values.joinToString(":"))
        }
    "#);
    assert_eq!(out, &["3", "a:b:c"]);
}

#[test]
fn test_list_of_not_null_removes_null_arguments() {
    let out = run_prints(r#"
        fun main() {
            val values = listOfNotNull(1, null, 2, null, 3)
            println(values.size)
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["3", "1,2,3"]);
}

#[test]
fn test_list_of_not_null_all_non_null_kept() {
    let out = run_prints(r#"
        fun main() {
            val values = listOfNotNull("a", "b", "c")
            println(values.size)
            println(values.joinToString(""))
        }
    "#);
    assert_eq!(out, &["3", "abc"]);
}

#[test]
fn test_list_of_not_null_from_expressions() {
    let out = run_prints(r#"
        fun main() {
            val values = listOfNotNull(if (true) "on" else null, null, if (false) "no" else "yes")
            println(values.joinToString("-"))
        }
    "#);
    assert_eq!(out, &["on-yes"]);
}

#[test]
fn test_empty_list_of_not_null() {
    let out = run_prints(r#"
        fun main() {
            val values = listOfNotNull<Int>()
            println(values.isEmpty())
            println(values.size)
        }
    "#);
    assert_eq!(out, &["true", "0"]);
}

#[test]
fn test_mutable_list_add_and_remove() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableListOf(1, 2)
            values.add(3)
            values.removeAt(1)
            values.add(1, 8)
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,8,3"]);
}

#[test]
fn test_mutable_list_clear_resets_size() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableListOf("a", "b", "c")
            values.clear()
            println(values.isEmpty())
            values.add("x")
            println(values.size)
        }
    "#);
    assert_eq!(out, &["true", "1"]);
}

#[test]
fn test_array_list_is_mutable_collection() {
    let out = run_prints(r#"
        fun main() {
            val values = ArrayList<Int>()
            values.add(9)
            values.add(7)
            values.add(5)
            values[1] = 4
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["9,4,5"]);
}

#[test]
fn test_array_list_insert_at_index() {
    let out = run_prints(r#"
        fun main() {
            val values = arrayListOf("x", "z")
            values.add(1, "y")
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["x,y,z"]);
}

#[test]
fn test_set_of_keeps_unique_elements() {
    let out = run_prints(r#"
        fun main() {
            val values = setOf(1, 2, 2, 3, 3, 1)
            println(values.size)
            println(values.contains(2))
            println(values.contains(4))
        }
    "#);
    assert_eq!(out, &["3", "true", "false"]);
}

#[test]
fn test_mutable_set_add_duplicate_is_noop() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableSetOf("a")
            val first = values.add("a")
            val second = values.add("b")
            println(first)
            println(second)
            println(values.size)
        }
    "#);
    assert_eq!(out, &["false", "true", "2"]);
}

#[test]
fn test_build_set_filters_duplicates() {
    let out = run_prints(r#"
        fun main() {
            val values = buildSet {
                add(1)
                add(2)
                add(2)
                add(3)
            }
            println(values.size)
            println(values.contains(2))
            println(values.contains(9))
        }
    "#);
    assert_eq!(out, &["3", "true", "false"]);
}

#[test]
fn test_linked_set_preserves_insertion_order_for_distinct_values() {
    let out = run_prints(r#"
        fun main() {
            val values = linkedSetOf("z", "a", "m")
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["z,a,m"]);
}

#[test]
fn test_sorted_set_orders_lexicographically() {
    let out = run_prints(r#"
        fun main() {
            val values = sortedSetOf(5, 1, 4, 2, 3)
            println(values.joinToString(","))
            println(values.first())
            println(values.last())
        }
    "#);
    assert_eq!(out, &["1,2,3,4,5", "1", "5"]);
}

#[test]
fn test_build_set_does_not_depend_on_add_order_for_set_semantics() {
    let out = run_prints(r#"
        fun main() {
            val values = buildSet {
                add("b")
                add("a")
                add("a")
                add("c")
            }
            println(values.contains("a"))
            println(values.contains("c"))
            println(values.size)
        }
    "#);
    assert_eq!(out, &["true", "true", "3"]);
}

#[test]
fn test_map_of_single_entry_lookup() {
    let out = run_prints(r#"
        fun main() {
            val value = mapOf("a" to 1)
            println(value["a"])
            println(value["b"] ?: -1)
        }
    "#);
    assert_eq!(out, &["1", "-1"]);
}

#[test]
fn test_mutable_map_add_update_remove() {
    let out = run_prints(r#"
        fun main() {
            val value = mutableMapOf("x" to 1)
            value["y"] = 2
            value["x"] = 9
            value.remove("y")
            println(value["x"])
            println(value.containsKey("y"))
            println(value.size)
        }
    "#);
    assert_eq!(out, &["9", "false", "1"]);
}

#[test]
fn test_linked_map_preserves_insertion_order() {
    let out = run_prints(r#"
        fun main() {
            val value = linkedMapOf("first" to 1, "second" to 2, "third" to 3)
            println(value.keys.joinToString(","))
            value["second"] = 7
            println(value.keys.joinToString(","))
        }
    "#);
    assert_eq!(out, &["first,second,third", "first,second,third"]);
}

#[test]
fn test_build_map_from_pairs() {
    let out = run_prints(r#"
        fun main() {
            val value = buildMap {
                put(1, "a")
                put(2, "b")
                put(1, "c")
            }
            println(value.size)
            println(value[1])
            println(value[2])
        }
    "#);
    assert_eq!(out, &["2", "c", "b"]);
}

#[test]
fn test_hash_map_allows_null_value_storage() {
    let out = run_prints(r#"
        fun main() {
            val value = HashMap<String, String?>()
            value["a"] = null
            value["b"] = "ok"
            println(value["a"] == null)
            println(value["b"])
            println(value.size)
        }
    "#);
    assert_eq!(out, &["true", "ok", "2"]);
}

#[test]
fn test_hash_map_contains_checks_both_key_and_value() {
    let out = run_prints(r#"
        fun main() {
            val value = hashMapOf("x" to 10, "y" to 20)
            println(value.containsKey("x"))
            println(value.containsValue(20))
            println(value.containsValue(30))
        }
    "#);
    assert_eq!(out, &["true", "true", "false"]);
}

#[test]
fn test_associate_from_sequence() {
    let out = run_prints(r#"
        fun main() {
            val values = sequenceOf("a", "bb", "ccc").associateWith { it.length }
            println(values["a"])
            println(values["ccc"])
        }
    "#);
    assert_eq!(out, &["1", "3"]);
}

#[test]
fn test_associate_with_merge_by_key() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf("ax", "ay", "bx").associate { it[0].toString() to it }
            println(values["a"])
            println(values["b"])
        }
    "#);
    assert_eq!(out, &["ay", "bx"]);
}

#[test]
fn test_group_by_with_map_factory() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(1, 2, 3, 4).groupBy { it % 2 == 0 }
            val even = values[true]?.sorted() ?: emptyList()
            val odd = values[false]?.sorted() ?: emptyList()
            println(even.joinToString(","))
            println(odd.joinToString(","))
        }
    "#);
    assert_eq!(out, &["2,4", "1,3"]);
}

#[test]
fn test_flatten_with_map_of_lists() {
    let out = run_prints(r#"
        fun main() {
            val value = mapOf("a" to listOf(1, 2), "b" to listOf(3), "c" to emptyList())
            val flat = value.values.flatten()
            println(flat.joinToString(","))
            println(flat.size)
        }
    "#);
    assert_eq!(out, &["1,2,3", "3"]);
}

#[test]
fn test_reified_builders_do_not_escape_mutation() {
    let out = run_prints(r#"
        fun main() {
            val base = mutableListOf(1)
            val built = buildList {
                addAll(base)
                add(2)
            }
            base.add(9)
            println(base.joinToString(","))
            println(built.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,9", "1,2"]);
}
