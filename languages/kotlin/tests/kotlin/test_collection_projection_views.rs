use crate::helpers::run_prints;

#[test]
fn test_map_key_set_view_reflects_updates() {
    let out = run_prints(
        r#"
        fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            val keys = map.keys
            val values = map.values
            println(keys.joinToString(","))
            println(values.joinToString(","))
            map["c"] = 3
            map["a"] = 9
            println(keys.joinToString(","))
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["a,b", "1,2", "a,b,c", "9,2,3"]);
}

#[test]
fn test_map_entry_set_mutation() {
    let out = run_prints(
        r#"
        fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            val entries = map.entries
            val it = entries.iterator()
            while (it.hasNext()) {
                val e = it.next()
                if (e.key == "a") {
                    it.remove()
                }
            }
            println(map.size)
            println(map.containsKey("a"))
            println(map["b"])
        }
    "#,
    );
    assert_eq!(out, &["1", "false", "2"]);
}

#[test]
fn test_map_values_mutable_list_backed_view() {
    let out = run_prints(
        r#"
        fun main() {
            val map = linkedMapOf("x" to 1, "y" to 2)
            val values = map.values
            println(values.sum())
            map["x"] = 9
            println(values.joinToString(","))
            val copied = values.toMutableList()
            copied.remove(2)
            println(values.size)
            println(copied.size)
        }
    "#,
    );
    assert_eq!(out, &["3", "9,2", "2", "1"]);
}

#[test]
fn test_list_as_view_from_sequence_to_list_projection() {
    let out = run_prints(
        r#"
        fun main() {
            val seq = sequenceOf(1, 2, 3, 4)
            val view = seq.toList()
            val rev = view.asReversed()
            println(rev.joinToString(","))
            println(view[0])
            println(view.last())
        }
    "#,
    );
    assert_eq!(out, &["4,3,2,1", "1", "4"]);
}

#[test]
fn test_iterator_mutability_contracts() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2, 3)
            val it = values.listIterator()
            println(it.next())
            it.set(7)
            println(values.joinToString(","))
            it.add(9)
            println(values.joinToString(","))
            it.previous()
            it.remove()
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1", "7,2,3", "7,9,2,3", "7,2,3"]);
}

#[test]
fn test_distinct_and_union_projection() {
    let out = run_prints(
        r#"
        fun main() {
            val a = listOf(1, 2, 2, 3)
            val b = listOf(2, 3, 4)
            val uniq = a.distinct()
            val union = a.union(b)
            println(uniq.joinToString(","))
            println(union.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "1,2,2,3,4"]);
}

#[test]
fn test_intersect_retains_order_in_list_result() {
    let out = run_prints(
        r#"
        fun main() {
            val a = listOf(1, 3, 5, 7, 3)
            val b = listOf(3, 3, 7)
            val out1 = a.intersect(b.toSet())
            println(out1.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["3,7"]);
}

#[test]
fn test_chunked_sliding_windowing_views() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(1, 2, 3, 4, 5)
            println(values.chunked(2).joinToString("|") { it.joinToString(":") })
            println(values.windowed(3).joinToString("|") { it.joinToString(":") })
            println(values.windowed(3, partialWindows = true).joinToString("|") { it.joinToString(":") })
        }
    "#,
    );
    assert_eq!(
        out,
        &["1:2|3:4|5", "1:2:3|2:3:4|3:4:5", "1:2:3|2:3:4|3:4:5|4:5|5"]
    );
}

#[test]
fn test_flatten_nested_projection() {
    let out = run_prints(
        r#"
        fun main() {
            val nested = listOf(listOf(1, 2), listOf(3), listOf(4, 5))
            println(nested.flatten().joinToString(","))
            println(nested.flatMap { it.map { v -> v * 2 } }.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3,4,5", "2,4,6,8,10"]);
}

#[test]
fn test_map_not_null_projection() {
    let out = run_prints(
        r#"
        fun main() {
            val source: List<Int?> = listOf(1, null, 3, null, 5)
            println(source.filterNotNull().joinToString(","))
            println(source.mapNotNull { it?.plus(10) }.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,3,5", "11,13,15"]);
}

#[test]
fn test_partition_and_grouping_projection() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(1, 2, 3, 4, 5)
            val (even, odd) = values.partition { it % 2 == 0 }
            println(even.joinToString(","))
            println(odd.joinToString(","))
            val byMod = values.groupBy { it % 2 }
            println(byMod[0]!!.joinToString(","))
            println(byMod[1]!!.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,4", "1,3,5", "2,4", "1,3,5"]);
}
