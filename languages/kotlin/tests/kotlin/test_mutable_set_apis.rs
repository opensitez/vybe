use crate::helpers::run_prints;

#[test]
fn test_mutable_set_add_and_size() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf(1, 2)
            values.add(3)
            println(values.size)
            println(values.contains(3))
        }
    "#,
    );
    assert_eq!(out, &["3", "true"]);
}

#[test]
fn test_mutable_set_add_duplicate_ignored() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf(1, 2)
            val added = values.add(2)
            println(added)
            println(values.size)
        }
    "#,
    );
    assert_eq!(out, &["false", "2"]);
}

#[test]
fn test_mutable_set_remove_existing() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3)
            val removed = values.remove(2)
            println(removed)
            println(values.size)
        }
    "#,
    );
    assert_eq!(out, &["true", "2"]);
}

#[test]
fn test_mutable_set_clear_empty() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3)
            values.clear()
            println(values.isEmpty())
            println(values.size)
        }
    "#,
    );
    assert_eq!(out, &["true", "0"]);
}

#[test]
fn test_mutable_set_union_via_plus_assign() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf(1, 2)
            values += 3
            values += 2
            println(values.joinToString(","))
            println(values.size)
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "3"]);
}

#[test]
fn test_mutable_set_minus_assign_removes() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3)
            values -= 2
            println(values.joinToString(","))
            println(values.size)
        }
    "#,
    );
    assert_eq!(out, &["1,3", "2"]);
}

#[test]
fn test_mutable_set_retain_all() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3, 4)
            values.retainAll(listOf(2, 4, 6))
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,4"]);
}

#[test]
fn test_mutable_set_remove_all() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3, 4)
            values.removeAll(listOf(1, 4))
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,3"]);
}

#[test]
fn test_mutable_set_contains_none() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf("a", "b")
            println(values.contains("c"))
            println(values.containsAll(listOf("a", "b")))
        }
    "#,
    );
    assert_eq!(out, &["false", "true"]);
}

#[test]
fn test_mutable_set_filtering_behavior() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3, 4)
            val filtered = values.filter { it > 2 }.toMutableSet()
            println(filtered.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["3,4"]);
}

#[test]
fn test_mutable_set_intersect_sets() {
    let out = run_prints(
        r#"
        fun main() {
            val a = mutableSetOf(1, 2, 3)
            val b = mutableSetOf(2, 3, 4)
            val c = a.intersect(b)
            println(c.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,3"]);
}

#[test]
fn test_mutable_set_union_sets() {
    let out = run_prints(
        r#"
        fun main() {
            val a = mutableSetOf(1, 2)
            val b = mutableSetOf(2, 3)
            val c = a.union(b)
            println(c.joinToString(","))
            println(c.size)
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "3"]);
}

#[test]
fn test_mutable_set_subtract() {
    let out = run_prints(
        r#"
        fun main() {
            val a = mutableSetOf(1, 2, 3)
            val b = mutableSetOf(2, 4)
            val c = a - b
            println(c.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,3"]);
}

#[test]
fn test_mutable_set_as_mutable_list() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3)
            val list = values.toMutableList()
            list.sort()
            println(list.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3"]);
}

#[test]
fn test_mutable_set_iterator_order_not_guaranteed_behavior() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf(3, 1, 2)
            val items = values.toMutableList()
            items.sort()
            println(items.joinToString(","))
            println(items.size)
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "3"]);
}

#[test]
fn test_mutable_set_first_last_like_functions() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf(2, 4, 6)
            println(values.first())
            println(values.last())
        }
    "#,
    );
    assert_eq!(out, &["2", "6"]);
}

#[test]
fn test_mutable_set_remove_if_even_not_present() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf(1, 3, 5)
            values.removeIf { it % 2 == 0 }
            println(values.joinToString(","))
            println(values.size)
        }
    "#,
    );
    assert_eq!(out, &["1,3,5", "3"]);
}

#[test]
fn test_mutable_set_add_all_from_list() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf(1)
            values.addAll(listOf(1, 2, 3))
            println(values.joinToString(","))
            println(values.size)
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "3"]);
}

#[test]
fn test_mutable_set_keep_distinct_order_after_mutation() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf(1, 2)
            values.add(2)
            values.add(3)
            values.remove(1)
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,3"]);
}

#[test]
fn test_mutable_set_filter_by_string_prefix() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf("aa", "ab", "bb")
            val filtered = values.filter { it.startsWith("a") }.toMutableSet()
            println(filtered.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["aa,ab"]);
}

#[test]
fn test_mutable_set_is_not_empty() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf<Int>()
            println(values.isEmpty())
            values.add(1)
            println(values.isNotEmpty())
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_mutable_set_contains_all() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3)
            println(values.containsAll(listOf(1, 3)))
            println(values.containsAll(listOf(1, 4)))
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_mutable_set_reassigned_through_to_mutable() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf(1, 2)
            val copy = values.toMutableSet()
            copy.add(3)
            println(values.joinToString(","))
            println(copy.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2", "1,2,3"]);
}

#[test]
fn test_mutable_set_join_to_string_delim() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3)
            println(values.joinToString("|"))
            println(values.joinToString(";") { (it * 2).toString() })
        }
    "#,
    );
    assert_eq!(out, &["1|2|3", "2;4;6"]);
}

#[test]
fn test_mutable_set_any_all_none() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableSetOf(1, 2, 3)
            println(values.any { it > 2 })
            println(values.all { it > 0 })
            println(values.none { it > 4 })
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true"]);
}
