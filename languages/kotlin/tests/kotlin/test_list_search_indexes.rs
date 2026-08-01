use crate::helpers::run_prints;

#[test]
fn test_binary_search_on_sorted_int_list() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(1, 3, 5, 7, 9)
            println(values.binarySearch(5))
            println(values.binarySearch(6))
        }
    "#,
    );
    assert_eq!(out, &["2", "-3"]);
}

#[test]
fn test_binary_search_with_range_window() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(1, 3, 5, 7, 9)
            println(values.binarySearch(5, 1, 4))
            println(values.binarySearch(5, 0, 2))
        }
    "#,
    );
    assert_eq!(out, &["2", "-2"]);
}

#[test]
fn test_binary_search_comparator_ordered_strings() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf("aa", "bb", "cc")
            val byLength = compareBy<String> { it.length }
            println(values.binarySearch("zz", comparator = byLength))
            println(values.binarySearch("b", comparator = byLength))
        }
    "#,
    );
    assert_eq!(out, &["-4", "0"]);
}

#[test]
fn test_index_of_first_predicate() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(2, 4, 6, 7, 8)
            println(values.indexOfFirst { it % 2 == 1 })
            println(values.indexOfFirst { it > 5 })
        }
    "#,
    );
    assert_eq!(out, &["3", "2"]);
}

#[test]
fn test_index_of_last_predicate() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(2, 4, 6, 7, 8)
            println(values.indexOfLast { it % 2 == 0 })
            println(values.indexOfLast { it > 10 })
        }
    "#,
    );
    assert_eq!(out, &["4", "-1"]);
}

#[test]
fn test_last_index_for_duplicates_and_absent_value() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(1, 4, 4, 4, 9)
            println(values.lastIndexOf(4))
            println(values.lastIndexOf(2))
        }
    "#,
    );
    assert_eq!(out, &["3", "-1"]);
}

#[test]
fn test_find_index_by_contains_check() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf("kotlin", "java", "python")
            val found = values.indexOfFirst { it.startsWith("jav") }
            val missing = values.indexOfFirst { it.startsWith("rust") }
            println(found)
            println(missing)
        }
    "#,
    );
    assert_eq!(out, &["1", "-1"]);
}

#[test]
fn test_index_of_range_and_stride() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(10, 20, 30, 40)
            println(values.indices.step(2).joinToString(","))
            println(values.slice(1..3).size)
            println(values.subList(1, 3).size)
        }
    "#,
    );
    assert_eq!(out, &["0,2", "3", "2"]);
}

#[test]
fn test_element_at_and_boundaries() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(11, 22, 33)
            println(values.elementAt(2))
            println(values.elementAtOrNull(5) ?: "none")
            println(values.elementAtOrElse(5) { value -> value * 10 + 1 })
        }
    "#,
    );
    assert_eq!(out, &["33", "none", "51"]);
}

#[test]
fn test_last_index_and_element_after_mutating_sublist() {
    let out = run_prints(
        r#"
        fun main() {
            val values = mutableListOf(1, 2, 3, 4)
            val window = values.subList(1, 3)
            println(window.lastIndex)
            window.clear()
            println(values.joinToString(","))
            println(values.lastIndex)
        }
    "#,
    );
    assert_eq!(out, &["1", "1,4", "1"]);
}

#[test]
fn test_list_contains_search_for_objects() {
    let out = run_prints(
        r#"
        fun main() {
            class Box(val v: Int)
            val values = listOf(Box(1), Box(2), Box(1))
            val a = Box(1)
            println(values.contains(a))
            println(values.indexOfFirst { it.v == 2 })
            println(values.indexOfLast { it.v == 1 })
        }
    "#,
    );
    assert_eq!(out, &["false", "1", "2"]);
}

#[test]
fn test_search_string_subsequence_patterns() {
    let out = run_prints(
        r#"
        fun main() {
            val lines = listOf("alpha", "beta", "gamma", "alphabet")
            println(lines.indexOf("beta"))
            println(lines.indexOfFirst { it.startsWith("alp") })
            println(lines.indexOfLast { it.endsWith("ta") })
        }
    "#,
    );
    assert_eq!(out, &["1", "0", "2"]);
}
