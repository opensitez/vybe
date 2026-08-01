use crate::helpers::run_prints;

#[test]
fn test_java_arrays_int_to_string() {
    let out = run_prints(
        r#"
        fun main() {
            val data = intArrayOf(5, 1, 4, 3, 2)
            println(java.util.Arrays.toString(data))
        }
    "#,
    );
    assert_eq!(out, &["[5, 1, 4, 3, 2]"]);
}

#[test]
fn test_java_arrays_int_sort_full() {
    let out = run_prints(
        r#"
        fun main() {
            val data = intArrayOf(5, 1, 4, 3, 2)
            java.util.Arrays.sort(data)
            println(java.util.Arrays.toString(data))
        }
    "#,
    );
    assert_eq!(out, &["[1, 2, 3, 4, 5]"]);
}

#[test]
fn test_java_arrays_int_sort_range_only() {
    let out = run_prints(
        r#"
        fun main() {
            val data = intArrayOf(9, 8, 7, 6, 5, 4)
            java.util.Arrays.sort(data, 1, 4)
            println(java.util.Arrays.toString(data))
        }
    "#,
    );
    assert_eq!(out, &["[9, 6, 7, 8, 5, 4]"]);
}

#[test]
fn test_java_arrays_int_binary_search_hit() {
    let out = run_prints(
        r#"
        fun main() {
            val data = intArrayOf(1, 2, 3, 4, 5)
            println(java.util.Arrays.binarySearch(data, 4))
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_java_arrays_int_binary_search_miss() {
    let out = run_prints(
        r#"
        fun main() {
            val data = intArrayOf(1, 2, 3, 4, 5)
            println(java.util.Arrays.binarySearch(data, 4))
            println(java.util.Arrays.binarySearch(data, 6))
        }
    "#,
    );
    assert_eq!(out, &["3", "-6"]);
}

#[test]
fn test_java_arrays_int_binary_search_in_range() {
    let out = run_prints(
        r#"
        fun main() {
            val data = intArrayOf(10, 20, 30, 40, 50)
            println(java.util.Arrays.binarySearch(data, 1, 4, 30))
            println(java.util.Arrays.binarySearch(data, 1, 4, 50))
        }
    "#,
    );
    assert_eq!(out, &["2", "-4"]);
}

#[test]
fn test_java_arrays_int_fill_all() {
    let out = run_prints(
        r#"
        fun main() {
            val data = intArrayOf(1, 2, 3, 4)
            java.util.Arrays.fill(data, 9)
            println(java.util.Arrays.toString(data))
        }
    "#,
    );
    assert_eq!(out, &["[9, 9, 9, 9]"]);
}

#[test]
fn test_java_arrays_int_fill_range() {
    let out = run_prints(
        r#"
        fun main() {
            val data = intArrayOf(0, 0, 0, 0, 0)
            java.util.Arrays.fill(data, 1, 4, 7)
            println(java.util.Arrays.toString(data))
        }
    "#,
    );
    assert_eq!(out, &["[0, 7, 7, 7, 0]"]);
}

#[test]
fn test_java_arrays_int_copy_of_extend() {
    let out = run_prints(
        r#"
        fun main() {
            val data = intArrayOf(1, 2, 3)
            val extended = java.util.Arrays.copyOf(data, 5)
            println(java.util.Arrays.toString(extended))
        }
    "#,
    );
    assert_eq!(out, &["[1, 2, 3, 0, 0]"]);
}

#[test]
fn test_java_arrays_int_copy_of_shrink() {
    let out = run_prints(
        r#"
        fun main() {
            val data = intArrayOf(1, 2, 3)
            val shrunk = java.util.Arrays.copyOf(data, 2)
            println(java.util.Arrays.toString(shrunk))
        }
    "#,
    );
    assert_eq!(out, &["[1, 2]"]);
}

#[test]
fn test_java_arrays_int_copy_of_range_middle() {
    let out = run_prints(
        r#"
        fun main() {
            val data = intArrayOf(1, 2, 3, 4, 5)
            val segment = java.util.Arrays.copyOfRange(data, 1, 4)
            println(java.util.Arrays.toString(segment))
        }
    "#,
    );
    assert_eq!(out, &["[2, 3, 4]"]);
}

#[test]
fn test_java_arrays_int_copy_of_range_empty() {
    let out = run_prints(
        r#"
        fun main() {
            val data = intArrayOf(1, 2, 3)
            val empty = java.util.Arrays.copyOfRange(data, 2, 2)
            println(java.util.Arrays.toString(empty))
        }
    "#,
    );
    assert_eq!(out, &["[]"]);
}

#[test]
fn test_java_arrays_int_hash_code() {
    let out = run_prints(
        r#"
        fun main() {
            val data = intArrayOf(1, 2, 3)
            println(java.util.Arrays.hashCode(data))
        }
    "#,
    );
    assert_eq!(out, &["30817"]);
}

#[test]
fn test_java_arrays_int_equality_edges() {
    let out = run_prints(
        r#"
        fun main() {
            val a = intArrayOf(1, 2, 3)
            val b = intArrayOf(1, 2, 3)
            val c = intArrayOf(1, 2, 4)
            println(java.util.Arrays.equals(a, b))
            println(java.util.Arrays.equals(a, c))
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_java_arrays_boolean_fill_and_string() {
    let out = run_prints(
        r#"
        fun main() {
            val flags = booleanArrayOf(true, false, true, false)
            java.util.Arrays.fill(flags, 1, 4, true)
            println(java.util.Arrays.toString(flags))
        }
    "#,
    );
    assert_eq!(out, &["[true, true, true, true]"]);
}

#[test]
fn test_java_arrays_char_fill_and_sort() {
    let out = run_prints(
        r#"
        fun main() {
            val data = charArrayOf('z', 'a', 'c', 'b')
            java.util.Arrays.sort(data)
            java.util.Arrays.fill(data, 1, 3, 'x')
            println(java.util.Arrays.toString(data))
        }
    "#,
    );
    assert_eq!(out, &["[a, x, x, z]"]);
}

#[test]
fn test_java_arrays_string_to_string_and_sort() {
    let out = run_prints(
        r#"
        fun main() {
            val data = arrayOf("k", "a", "m", "b")
            java.util.Arrays.sort(data)
            println(java.util.Arrays.toString(data))
        }
    "#,
    );
    assert_eq!(out, &["[a, b, k, m]"]);
}

#[test]
fn test_java_arrays_string_fill() {
    let out = run_prints(
        r#"
        fun main() {
            val data = arrayOf("left", "right", "up")
            java.util.Arrays.fill(data, 1, 2, "mid")
            println(java.util.Arrays.toString(data))
        }
    "#,
    );
    assert_eq!(out, &["[left, mid, up]"]);
}

#[test]
fn test_java_arrays_string_copy_of_and_mutability_gap() {
    let out = run_prints(
        r#"
        fun main() {
            val source = arrayOf("x", "y", "z")
            val copy = java.util.Arrays.copyOf(source, 4)
            copy[1] = "changed"
            println(source[1])
            println(java.util.Arrays.toString(copy))
        }
    "#,
    );
    assert_eq!(out, &["y", "[x, changed, z, null]"]);
}

#[test]
fn test_java_arrays_string_copy_of_range() {
    let out = run_prints(
        r#"
        fun main() {
            val source = arrayOf("a", "b", "c", "d")
            val segment = java.util.Arrays.copyOfRange(source, 2, 4)
            println(java.util.Arrays.toString(segment))
        }
    "#,
    );
    assert_eq!(out, &["[c, d]"]);
}

#[test]
fn test_java_arrays_as_list_backing_array() {
    let out = run_prints(
        r#"
        fun main() {
            val data = arrayOf("one", "two", "three")
            val view = java.util.Arrays.asList(data)
            view[1] = "changed"
            println(data[1])
            println(view.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["changed", "one,changed,three"]);
}

#[test]
fn test_java_arrays_as_list_search_and_size() {
    let out = run_prints(
        r#"
        fun main() {
            val data = arrayOf("a", "b", "c")
            val view = java.util.Arrays.asList(data)
            println(view.size)
            println(view.indexOf("b"))
            println(view.contains("c"))
        }
    "#,
    );
    assert_eq!(out, &["3", "1", "true"]);
}

#[test]
fn test_java_arrays_string_sort_with_comparator_desc() {
    let out = run_prints(
        r#"
        fun main() {
            val data = arrayOf("aa", "b", "cccc", "ddd")
            java.util.Arrays.sort(data, java.util.Comparator { a, b ->
                b.length - a.length
            })
            println(java.util.Arrays.toString(data))
        }
    "#,
    );
    assert_eq!(out, &["[cccc, ddd, aa, b]"]);
}

#[test]
fn test_java_arrays_string_binary_search_hit_and_miss() {
    let out = run_prints(
        r#"
        fun main() {
            val data = arrayOf("a", "b", "m", "z")
            java.util.Arrays.sort(data)
            println(java.util.Arrays.binarySearch(data, "m"))
            println(java.util.Arrays.binarySearch(data, "k"))
        }
    "#,
    );
    assert_eq!(out, &["2", "-3"]);
}

#[test]
fn test_java_arrays_deep_to_string_nested() {
    let out = run_prints(
        r#"
        fun main() {
            val nested = arrayOf(arrayOf(1, 2), arrayOf(3, 4))
            println(java.util.Arrays.deepToString(nested))
        }
    "#,
    );
    assert_eq!(out, &["[[1, 2], [3, 4]]"]);
}

#[test]
fn test_java_arrays_deep_equals_nested_match_and_miss() {
    let out = run_prints(
        r#"
        fun main() {
            val lhs = arrayOf(arrayOf(1, 2), arrayOf(3, 4))
            val rhs = arrayOf(arrayOf(1, 2), arrayOf(3, 4))
            val other = arrayOf(arrayOf(1, 2), arrayOf(3, 5))
            println(java.util.Arrays.deepEquals(lhs, rhs))
            println(java.util.Arrays.deepEquals(lhs, other))
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_java_arrays_deep_hash_code_is_stable_for_same_structure() {
    let out = run_prints(
        r#"
        fun main() {
            val lhs = arrayOf(arrayOf(1, 2), arrayOf(3, 4))
            val rhs = arrayOf(arrayOf(1, 2), arrayOf(3, 4))
            println(java.util.Arrays.deepHashCode(lhs) == java.util.Arrays.deepHashCode(rhs))
        }
    "#,
    );
    assert_eq!(out, &["true"]);
}

#[test]
fn test_java_arrays_long_copy_of_range_and_sort() {
    let out = run_prints(
        r#"
        fun main() {
            val data = longArrayOf(9L, 4L, 7L, 1L, 8L, 2L)
            val segment = java.util.Arrays.copyOfRange(data, 1, 5)
            java.util.Arrays.sort(segment)
            println(java.util.Arrays.toString(segment))
        }
    "#,
    );
    assert_eq!(out, &["[1, 4, 7, 8]"]);
}
