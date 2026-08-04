use crate::helpers::run_prints;

#[test]
fn test_map_keys_projection() {
    let out = run_prints(
        r#"
        fun main() {
            val source = mapOf("a" to 1, "bb" to 2)
            val upper = source.mapKeys { it.key.uppercase() }
            println(upper.keys.joinToString(","))
            println(upper["A"])
        }
    "#,
    );
    assert_eq!(out, &["A,BB", "1"]);
}

#[test]
fn test_map_values_projection() {
    let out = run_prints(
        r#"
        fun main() {
            val source = mapOf("a" to 1, "b" to 2)
            val doubled = source.mapValues { it.value * 2 }
            println(doubled.values.joinToString(","))
            println(doubled["b"])
        }
    "#,
    );
    assert_eq!(out, &["2,4", "4"]);
}

#[test]
fn test_map_filter_keys() {
    let out = run_prints(
        r#"
        fun main() {
            val source = mapOf("a" to 1, "bb" to 2, "ccc" to 3)
            val short = source.filterKeys { it.length <= 2 }
            println(short.size)
            println(short.keys.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2", "a,bb"]);
}

#[test]
fn test_map_filter_values() {
    let out = run_prints(
        r#"
        fun main() {
            val source = mapOf("x" to 3, "y" to 12, "z" to 18)
            val filtered = source.filterValues { it % 6 == 0 }
            println(filtered.keys.joinToString(","))
            println(filtered["z"])
        }
    "#,
    );
    assert_eq!(out, &["y,z", "18"]);
}

#[test]
fn test_map_filter_pairs_combined() {
    let out = run_prints(
        r#"
        fun main() {
            val source = mapOf("aa" to 1, "bb" to 2, "ac" to 3)
            val projected = source.filter { it.key.startsWith("a") && it.value > 1 }
            println(projected.size)
            println(projected["ac"])
        }
    "#,
    );
    assert_eq!(out, &["1", "3"]);
}

#[test]
fn test_map_get_or_put_build_lists() {
    let out = run_prints(
        r#"
        fun main() {
            val buckets = linkedMapOf<String, MutableList<Int>>()
            buckets.getOrPut("a") { mutableListOf() }.add(1)
            buckets.getOrPut("a") { mutableListOf() }.add(2)
            buckets.getOrPut("b") { mutableListOf() }.add(9)
            println(buckets["a"]!!.joinToString(","))
            println(buckets.size)
        }
    "#,
    );
    assert_eq!(out, &["1,2", "2"]);
}

#[test]
fn test_map_put_if_absent_updates_once() {
    let out = run_prints(
        r#"
        fun main() {
            val source = linkedMapOf("a" to 1)
            val existing = source.putIfAbsent("a", 99)
            val added = source.putIfAbsent("b", 2)
            println(existing)
            println(added)
            println(source["a"])
            println(source["b"])
        }
    "#,
    );
    assert_eq!(out, &["1", "null", "1", "2"]);
}

#[test]
fn test_map_get_or_default_and_else() {
    let out = run_prints(
        r#"
        fun main() {
            val source = mapOf("one" to 1, "two" to 2)
            println(source.getOrDefault("two", -1))
            println(source.getOrDefault("x", -1))
            println(source.getOrElse("two") { 0 })
            println(source.getOrElse("x") { 7 })
        }
    "#,
    );
    assert_eq!(out, &["2", "-1", "2", "7"]);
}

#[test]
fn test_map_get_value_and_or_null() {
    let out = run_prints(
        r#"
        fun main() {
            val source = mapOf("a" to 1)
            try {
                println(source.getValue("a"))
                println(source.getValue("b"))
            } catch (e: Exception) {
                println("err")
            }
        }
    "#,
    );
    assert_eq!(out, &["1", "err"]);
}

#[test]
fn test_map_to_pair_lists_round_trip() {
    let out = run_prints(
        r#"
        fun main() {
            val source = mapOf("x" to 9, "y" to 8)
            val pairs = source.toList()
            val rebuilt = pairs.toMap()
            println(pairs.joinToString("|") { it.toString() })
            println(rebuilt.size)
            println(rebuilt["y"])
        }
    "#,
    );
    assert_eq!(out, &["(x, 9)|(y, 8)", "2", "8"]);
}

#[test]
fn test_map_entries_projection_with_indexed_map() {
    let out = run_prints(
        r#"
        fun main() {
            val source = linkedMapOf("a" to 1, "b" to 2, "c" to 3)
            val indexed = source.entries.mapIndexed { index, e -> "${index}:${e.key}:${e.value}" }
            println(indexed.joinToString("|"))
        }
    "#,
    );
    assert_eq!(out, &["0:a:1|1:b:2|2:c:3"]);
}

#[test]
fn test_map_mutable_conversion_views() {
    let out = run_prints(
        r#"
        fun main() {
            val source = mapOf("a" to 1, "b" to 2)
            val mutable = source.toMutableMap()
            mutable["c"] = 3
            mutable.remove("a")
            println(mutable.size)
            println(mutable.keys.joinToString(","))
            val restored = mutable.toMap()
            println(restored.containsKey("a"))
            println(restored["c"])
        }
    "#,
    );
    assert_eq!(out, &["2", "b,c", "false", "3"]);
}
