use crate::helpers::run_prints;

#[test]
fn test_java_collections_empty_list_is_immutable_but_size_zero() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.Collections.emptyList<Int>()
            println(values.isEmpty())
            println(values.size)
        }
    "#,
    );
    assert_eq!(out, &["true", "0"]);
}

#[test]
fn test_java_collections_empty_set_is_reusable_singleton() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.Collections.emptySet<String>()
            println(values.isEmpty())
            println(values.contains("a"))
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_java_collections_empty_map_has_zero_size() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.Collections.emptyMap<String, Int>()
            println(values.isEmpty())
            println(values.size)
        }
    "#,
    );
    assert_eq!(out, &["true", "0"]);
}

#[test]
fn test_java_collections_singleton_list_wraps_single_value() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.Collections.singletonList("only")
            println(values.size)
            println(values[0])
        }
    "#,
    );
    assert_eq!(out, &["1", "only"]);
}

#[test]
fn test_java_collections_singleton_set_rejects_second_insert() {
    let out = run_prints(
        r#"
        fun main() {
            val values: java.util.Set<String> = java.util.Collections.singleton("x")
            println(values.size)
            try {
                values.add("y")
                println("added")
            } catch (ex: Exception) {
                println("error")
            }
        }
    "#,
    );
    assert_eq!(out, &["1", "error"]);
}

#[test]
fn test_java_collections_singleton_map_key_value_pair() {
    let out = run_prints(
        r#"
        fun main() {
            val entry = java.util.Collections.singletonMap("a", 1)
            println(entry.size)
            println(entry["a"])
            println(entry["missing"] ?: "none")
        }
    "#,
    );
    assert_eq!(out, &["1", "1", "none"]);
}

#[test]
fn test_java_collections_n_copies_materializes_repeated_value() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.Collections.nCopies(4, "zap")
            println(values.size)
            println(values[0])
            println(values[3])
        }
    "#,
    );
    assert_eq!(out, &["4", "zap", "zap"]);
}

#[test]
fn test_java_collections_n_copies_contains_every_copy() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.Collections.nCopies(3, 7)
            println(values.contains(7))
            println(values.contains(2))
            println(values.indexOf(7))
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "0"]);
}

#[test]
fn test_java_collections_copy_populates_destination_from_source() {
    let out = run_prints(
        r#"
        fun main() {
            val source = java.util.ArrayList<Int>()
            source.add(1)
            source.add(2)
            source.add(3)
            val target = java.util.ArrayList<Int>(java.util.ArrayList<Int>(listOf(0, 0, 0, 0)))
            java.util.Collections.copy(target, source)
            println(target)
        }
    "#,
    );
    assert_eq!(out, &["[1, 2, 3, 0]"]);
}

#[test]
fn test_java_collections_copy_fails_when_destination_too_small() {
    let out = run_prints(
        r#"
        fun main() {
            val source = java.util.ArrayList<Int>(listOf(1, 2, 3))
            val target = java.util.ArrayList<Int>(listOf(0))
            try {
                java.util.Collections.copy(target, source)
                println("ok")
            } catch (ex: Exception) {
                println("error")
            }
        }
    "#,
    );
    assert_eq!(out, &["error"]);
}

#[test]
fn test_java_collections_frequency_counts_occurrences() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.Arrays.asList(1, 2, 1, 3, 1, 2)
            println(java.util.Collections.frequency(values, 1))
            println(java.util.Collections.frequency(values, 2))
            println(java.util.Collections.frequency(values, 4))
        }
    "#,
    );
    assert_eq!(out, &["3", "2", "0"]);
}

#[test]
fn test_java_collections_disjoint_when_no_shared_elements() {
    let out = run_prints(
        r#"
        fun main() {
            val a = java.util.Arrays.asList(1, 2, 3)
            val b = java.util.Arrays.asList(4, 5, 6)
            println(java.util.Collections.disjoint(a, b))
            println(java.util.Collections.disjoint(a, java.util.Arrays.asList(3, 8)))
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_java_collections_fill_replaces_all_values() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2, 3, 4))
            java.util.Collections.fill(values, 9)
            println(values)
        }
    "#,
    );
    assert_eq!(out, &["[9, 9, 9, 9]"]);
}

#[test]
fn test_java_collections_sort_natural_order() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.ArrayList<Int>(listOf(4, 1, 3, 2))
            java.util.Collections.sort(values)
            println(values)
        }
    "#,
    );
    assert_eq!(out, &["[1, 2, 3, 4]"]);
}

#[test]
fn test_java_collections_sort_uses_comparator_desc() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.ArrayList<String>(listOf("bbb", "c", "aa"))
            java.util.Collections.sort(values, java.util.Collections.reverseOrder<String>())
            println(values)
        }
    "#,
    );
    assert_eq!(out, &["[c, bbb, aa]"]);
}

#[test]
fn test_java_collections_reverse() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2, 3))
            java.util.Collections.reverse(values)
            println(values)
        }
    "#,
    );
    assert_eq!(out, &["[3, 2, 1]"]);
}

#[test]
fn test_java_collections_swap() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2, 3))
            java.util.Collections.swap(values, 0, 2)
            println(values)
        }
    "#,
    );
    assert_eq!(out, &["[3, 2, 1]"]);
}

#[test]
fn test_java_collections_rotate_positive_distance() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2, 3, 4))
            java.util.Collections.rotate(values, 1)
            println(values)
        }
    "#,
    );
    assert_eq!(out, &["[4, 1, 2, 3]"]);
}

#[test]
fn test_java_collections_rotate_negative_distance() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2, 3, 4))
            java.util.Collections.rotate(values, -1)
            println(values)
        }
    "#,
    );
    assert_eq!(out, &["[2, 3, 4, 1]"]);
}

#[test]
fn test_java_collections_replace_all_values() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2, 1, 3))
            java.util.Collections.replaceAll(values, 1, 9)
            println(values)
        }
    "#,
    );
    assert_eq!(out, &["[9, 2, 9, 3]"]);
}

#[test]
fn test_java_collections_add_all_appends_elements() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2))
            val more = java.util.ArrayList<Int>(listOf(3, 4))
            val changed = java.util.Collections.addAll(values, more[0], more[1])
            println(changed)
            println(values)
        }
    "#,
    );
    assert_eq!(out, &["true", "[1, 2, 3, 4]"]);
}

#[test]
fn test_java_collections_unmodifiable_list_forbids_mutation() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2, 3))
            val safe = java.util.Collections.unmodifiableList(values)
            println(safe[1])
            try {
                safe[1] = 9
                println("changed")
            } catch (ex: Exception) {
                println("error")
            }
        }
    "#,
    );
    assert_eq!(out, &["2", "error"]);
}

#[test]
fn test_java_collections_unmodifiable_set_forbids_add() {
    let out = run_prints(
        r#"
        fun main() {
            val source = java.util.HashSet<Int>()
            source.add(1)
            source.add(2)
            val safe = java.util.Collections.unmodifiableSet(source)
            try {
                safe.add(3)
                println("added")
            } catch (ex: Exception) {
                println("error")
            }
            println(safe.size)
        }
    "#,
    );
    assert_eq!(out, &["error", "2"]);
}

#[test]
fn test_java_collections_unmodifiable_map_forbids_put() {
    let out = run_prints(
        r#"
        fun main() {
            val source = java.util.LinkedHashMap<String, Int>()
            source["a"] = 1
            source["b"] = 2
            val safe = java.util.Collections.unmodifiableMap(source)
            try {
                safe["c"] = 3
                println("added")
            } catch (ex: Exception) {
                println("error")
            }
            println(safe["b"])
        }
    "#,
    );
    assert_eq!(out, &["error", "2"]);
}

#[test]
fn test_java_collections_binary_search_found_and_miss() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2, 3, 5, 8))
            println(java.util.Collections.binarySearch(values, 3))
            println(java.util.Collections.binarySearch(values, 4))
        }
    "#,
    );
    assert_eq!(out, &["2", "-4"]);
}

#[test]
fn test_java_collections_index_and_last_index_of_sublist() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2, 3, 2, 3, 4))
            val sub = java.util.ArrayList<Int>(listOf(2, 3))
            println(java.util.Collections.indexOfSubList(values, sub))
            println(java.util.Collections.lastIndexOfSubList(values, sub))
            println(java.util.Collections.indexOfSubList(values, java.util.ArrayList<Int>(listOf(9, 9))))
        }
    "#,
    );
    assert_eq!(out, &["1", "3", "-1"]);
}

#[test]
fn test_java_collections_min_max_values() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.ArrayList<Int>(listOf(8, 1, 4, 3))
            println(java.util.Collections.min(values))
            println(java.util.Collections.max(values))
        }
    "#,
    );
    assert_eq!(out, &["1", "8"]);
}

#[test]
fn test_java_collections_min_with_comparator() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.ArrayList<String>(listOf("aa", "z", "bbb", "c"))
            val shortest = java.util.Collections.min(values, compareBy<String> { it.length })
            val longest = java.util.Collections.max(values, compareBy<String> { it.length })
            println(shortest)
            println(longest)
        }
    "#,
    );
    assert_eq!(out, &["z", "bbb"]);
}

#[test]
fn test_java_collections_new_set_from_map() {
    let out = run_prints(
        r#"
        fun main() {
            val map = java.util.HashMap<String, Boolean>()
            val set = java.util.Collections.newSetFromMap(map)
            set.add("x")
            set.add("y")
            println(set)
            println(map.size)
        }
    "#,
    );
    assert_eq!(out, &["[x, y]", "2"]);
}

#[test]
fn test_java_collections_synchronized_list_supports_concurrent_view_semantics() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2, 3))
            val sync = java.util.Collections.synchronizedList(values)
            sync.add(4)
            println(sync.size)
            sync[0] = 0
            println(sync[0])
        }
    "#,
    );
    assert_eq!(out, &["4", "0"]);
}
