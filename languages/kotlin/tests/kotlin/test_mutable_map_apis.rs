use crate::helpers::run_prints;

#[test]
fn test_mutable_map_put_and_get() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1)
            val prev = values.put("b", 2)
            println(prev)
            println(values["b"])
        }
    "#,
    );
    assert_eq!(out, &["null", "2"]);
}

#[test]
fn test_mutable_map_put_replaces_existing() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1)
            val prev = values.put("a", 9)
            println(prev)
            println(values["a"])
        }
    "#,
    );
    assert_eq!(out, &["1", "9"]);
}

#[test]
fn test_mutable_map_remove_key() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1, "b" to 2)
            val removed = values.remove("a")
            println(removed)
            println(values.size)
        }
    "#,
    );
    assert_eq!(out, &["1", "1"]);
}

#[test]
fn test_mutable_map_remove_missing_key() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1)
            val removed = values.remove("z")
            println(removed == null)
            println(values.size)
        }
    "#,
    );
    assert_eq!(out, &["true", "1"]);
}

#[test]
fn test_mutable_map_remove_key_value_pair() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1, "b" to 2)
            val removed = values.remove("a", 1)
            println(removed)
            println(values.size)
        }
    "#,
    );
    assert_eq!(out, &["true", "1"]);
}

#[test]
fn test_mutable_map_remove_key_value_mismatch() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1)
            val removed = values.remove("a", 9)
            println(removed)
            println(values.size)
            println(values["a"])
        }
    "#,
    );
    assert_eq!(out, &["false", "1", "1"]);
}

#[test]
fn test_mutable_map_contains_key_and_value() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1, "b" to 2)
            println(values.containsKey("a"))
            println(values.containsValue(2))
            println(values.containsValue(3))
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "false"]);
}

#[test]
fn test_mutable_map_get_or_default() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1)
            println(values.getOrDefault("a", 9))
            println(values.getOrDefault("b", 9))
        }
    "#,
    );
    assert_eq!(out, &["1", "9"]);
}

#[test]
fn test_mutable_map_get_or_else() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1)
            println(values.getOrElse("a") { 9 })
            println(values.getOrElse("b") { 9 })
        }
    "#,
    );
    assert_eq!(out, &["1", "9"]);
}

#[test]
fn test_mutable_map_get_value_throws_if_missing() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1)
            try {
                println(values.getValue("b"))
            } catch (e: NoSuchElementException) {
                println("missing")
            }
        }
    "#,
    );
    assert_eq!(out, &["missing"]);
}

#[test]
fn test_mutable_map_get_or_put_existing() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1)
            val value = values.getOrPut("a") { 99 }
            println(value)
            println(values["a"])
        }
    "#,
    );
    assert_eq!(out, &["1", "1"]);
}

#[test]
fn test_mutable_map_get_or_put_missing() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1)
            val value = values.getOrPut("b") { 8 }
            println(value)
            println(values["b"])
        }
    "#,
    );
    assert_eq!(out, &["8", "8"]);
}

#[test]
fn test_mutable_map_put_all_values() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1)
            values.putAll(mapOf("b" to 2, "c" to 3))
            println(values.keys.joinToString(","))
            println(values["b"])
            println(values["c"])
        }
    "#,
    );
    assert_eq!(out, &["a,b,c", "2", "3"]);
}

#[test]
fn test_mutable_map_filter_keys() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1, "b" to 2, "c" to 3)
            val filtered = values.filterValues { it > 1 }
            println(filtered.joinToString(",") { it.key + ":" + it.value })
        }
    "#,
    );
    assert_eq!(out, &["b:2,c:3"]);
}

#[test]
fn test_mutable_map_map_keys() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1, "b" to 2)
            val mapped = values.mapKeys { it.key + it.value.toString() }
            println(mapped.keys.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["a1,b2"]);
}

#[test]
fn test_mutable_map_update_existing_values() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1, "b" to 2)
            values["a"] = 4
            values.put("b", 5)
            println(values["a"])
            println(values["b"])
        }
    "#,
    );
    assert_eq!(out, &["4", "5"]);
}

#[test]
fn test_mutable_map_entries_iteration() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("x" to 1, "y" to 2)
            val out = values.entries.joinToString("|") { it.key + ":" + it.value }
            println(out)
        }
    "#,
    );
    assert_eq!(out, &["x:1|y:2"]);
}

#[test]
fn test_mutable_map_keys_iteration() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("x" to 1, "y" to 2)
            val joined = values.keys.joinToString(",")
            println(joined)
        }
    "#,
    );
    assert_eq!(out, &["x,y"]);
}

#[test]
fn test_mutable_map_values_sum() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1, "b" to 2)
            println(values.values.sum())
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_mutable_map_is_empty_checks() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf<String, Int>()
            println(values.isEmpty())
            values["a"] = 1
            println(values.isNotEmpty())
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_mutable_map_compute_if_absent_like_pattern() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1)
            if (!values.containsKey("b")) {
                values["b"] = values.size
            }
            println(values["b"])
            println(values["a"])
        }
    "#,
    );
    assert_eq!(out, &["1", "1"]);
}

#[test]
fn test_mutable_map_retain_all_keys() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1, "b" to 2, "c" to 3)
            values.entries.removeIf { it.key == "b" }
            println(values.keys.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["a,c"]);
}

#[test]
fn test_mutable_map_plus_and_minus_assign() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1)
            values += mapOf("b" to 2)
            values -= "a"
            println(values.keys.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["b"]);
}

#[test]
fn test_mutable_map_mutable_entries_copy() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1)
            val copy = values.toMutableMap()
            copy["a"] = 9
            println(values["a"])
            println(copy["a"])
        }
    "#,
    );
    assert_eq!(out, &["1", "9"]);
}

#[test]
fn test_mutable_map_from_linked_hash_map_type() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.LinkedHashMap<String, Int>()
            values["a"] = 1
            values["b"] = 2
            println(values["a"])
            println(values.keys.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1", "a,b"]);
}

#[test]
fn test_mutable_map_merge_like_update() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1)
            val next = values.getOrElse("a") { 0 } + 4
            values["a"] = next
            println(values["a"])
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_mutable_map_to_immutable_map() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableMapOf("a" to 1)
            val frozen = values.toMap()
            println(frozen["a"])
            values["a"] = 9
            println(frozen["a"])
            println(values["a"])
        }
    "#,
    );
    assert_eq!(out, &["1", "1", "9"]);
}
