use crate::helpers::run_prints;

#[test]
fn test_mutable_list_add_and_get() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2)
            values.add(3)
            println(values[2])
            println(values.size)
        }
    "#,
    );
    assert_eq!(out, &["3", "3"]);
}

#[test]
fn test_mutable_list_add_at_index() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 3, 4)
            values.add(1, 2)
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3,4"]);
}

#[test]
fn test_mutable_list_add_all_from_list() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1)
            values.addAll(listOf(2, 3))
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3"]);
}

#[test]
fn test_mutable_list_set_replaces_value() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2, 3)
            values[1] = 9
            println(values[1])
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["9", "1,9,3"]);
}

#[test]
fn test_mutable_list_remove_value() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2, 3, 2)
            val removed = values.remove(2)
            println(removed)
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["true", "1,3,2"]);
}

#[test]
fn test_mutable_list_remove_at_index() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(7, 8, 9)
            val removed = values.removeAt(1)
            println(removed)
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["8", "7,9"]);
}

#[test]
fn test_mutable_list_clear_empties() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2)
            values.clear()
            println(values.isEmpty())
            println(values.size)
        }
    "#,
    );
    assert_eq!(out, &["true", "0"]);
}

#[test]
fn test_mutable_list_sub_list_view_modification() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2, 3, 4)
            val window = values.subList(1, 3)
            window.clear()
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,4"]);
}

#[test]
fn test_mutable_list_remove_first() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(9, 1, 2)
            println(values.removeFirst())
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["9", "1,2"]);
}

#[test]
fn test_mutable_list_remove_last() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2, 9)
            println(values.removeLast())
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["9", "1,2"]);
}

#[test]
fn test_mutable_list_first_and_last() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(4, 5, 6)
            println(values.first())
            println(values.last())
        }
    "#,
    );
    assert_eq!(out, &["4", "6"]);
}

#[test]
fn test_mutable_list_index_of_and_last_index_of() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2, 3, 2)
            println(values.indexOf(2))
            println(values.lastIndexOf(2))
        }
    "#,
    );
    assert_eq!(out, &["1", "3"]);
}

#[test]
fn test_mutable_list_keep_first_k() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2, 3, 4)
            val head = values.take(2)
            println(head.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2"]);
}

#[test]
fn test_mutable_list_drop_k() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2, 3, 4)
            val tail = values.drop(2)
            println(tail.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["3,4"]);
}

#[test]
fn test_mutable_list_retain_only_matching() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2, 3, 4)
            values.retainAll { it % 2 == 0 }
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,4"]);
}

#[test]
fn test_mutable_list_remove_if_even() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2, 3, 4)
            values.removeIf { it % 2 == 1 }
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,4"]);
}

#[test]
fn test_mutable_list_sort_ascending() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(4, 1, 3, 2)
            values.sort()
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3,4"]);
}

#[test]
fn test_mutable_list_sort_descending() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(4, 1, 3, 2)
            values.sortDescending()
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["4,3,2,1"]);
}

#[test]
fn test_mutable_list_shuffle_is_not_stable_shape() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2, 3)
            val copy = values.toMutableList()
            copy.shuffle()
            println(copy.size)
            println(copy.contains(1))
            println(copy.contains(2))
            println(copy.contains(3))
        }
    "#,
    );
    assert_eq!(out, &["3", "true", "true", "true"]);
}

#[test]
fn test_mutable_list_reverse_in_place() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2, 3)
            values.reverse()
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["3,2,1"]);
}

#[test]
fn test_mutable_list_replace_all() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2, 3)
            values.replaceAll { it * 2 }
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,4,6"]);
}

#[test]
fn test_mutable_list_fill_value() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2, 3)
            values.fill(9)
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["9,9,9"]);
}

#[test]
fn test_mutable_list_plus_assign() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1)
            values += 2
            values += 3
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3"]);
}

#[test]
fn test_mutable_list_index_iteration() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2, 3)
            var acc = 0
            for ((index, value) in values.withIndex()) {
                acc += index + value
            }
            println(acc)
        }
    "#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_mutable_list_contains_checks() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf("a", "b", "c")
            println(values.contains("b"))
            println(values.isNotEmpty())
            println(values.none())
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "false"]);
}

#[test]
fn test_mutable_list_ensure_order_after_reattribute() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2, 3)
            val copy = values.toMutableList()
            copy[0] = 9
            println(values.joinToString(","))
            println(copy.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "9,2,3"]);
}
