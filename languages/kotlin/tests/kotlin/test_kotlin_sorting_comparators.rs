use crate::helpers::run_prints;

#[test]
fn test_sorted_numbers_default_order() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(4, 1, 3, 2)
            println(values.sorted().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3,4"]);
}

#[test]
fn test_sorted_descending_numbers() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(1, 4, 2, 3)
            println(values.sortedDescending().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["4,3,2,1"]);
}

#[test]
fn test_sorted_strings_by_length() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf("bbb", "a", "cccc")
            println(values.sortedBy { it.length }.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["a,bbb,cccc"]);
}

#[test]
fn test_sorted_by_descending_length() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf("bbb", "a", "cccc")
            println(values.sortedByDescending { it.length }.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["cccc,bbb,a"]);
}

#[test]
fn test_sorted_with_reverse_numeric_comparator() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(1, 3, 2, 5, 4)
            val out = values.sortedWith(compareByDescending { it })
            println(out.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["5,4,3,2,1"]);
}

#[test]
fn test_sorted_with_custom_tie_breaker() {
    let out = run_prints(
        r#"
        fun main() {
            data class Item(val first: Int, val second: String)
            val values = listOf(Item(2, "b"), Item(1, "c"), Item(2, "a"))
            val out = values.sortedWith(compareBy<Item> { it.first }.thenBy { it.second })
            println(out.joinToString(",") { "${'$'}{it.first}${'$'}{it.second}" })
        }
    "#,
    );
    assert_eq!(out, &["1c,2a,2b"]);
}

#[test]
fn test_sorted_with_then_by_secondary() {
    let out = run_prints(
        r#"
        fun main() {
            data class Item(val first: Int, val second: Int)
            val values = listOf(Item(1, 9), Item(1, 2), Item(2, 5))
            val out = values.sortedWith(compareBy<Item> { it.first }.thenByDescending { it.second })
            println(out.joinToString(",") { "${'$'}{it.first}-${'$'}{it.second}" })
        }
    "#,
    );
    assert_eq!(out, &["1-9,1-2,2-5"]);
}

#[test]
fn test_reversed_list() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(1, 2, 3, 4)
            println(values.reversed().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["4,3,2,1"]);
}

#[test]
fn test_reversed_of_strings() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf("a", "b", "c")
            println(values.reversed())
        }
    "#,
    );
    assert_eq!(out, &["[c, b, a]"]);
}

#[test]
fn test_as_reversed_views_original_mutation_reflected() {
    let out = run_prints(
        r#"
        fun main() {
            val mutable = mutableListOf(1, 2, 3)
            val reversed = mutable.asReversed()
            reversed[0] = 9
            println(mutable.joinToString(","))
            println(reversed.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["9,2,1", "1,2,9"]);
}

#[test]
fn test_min_of_returns_smallest() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(7, 3, 9, 1)
            println(values.minOrNull())
            println(values.maxOrNull())
        }
    "#,
    );
    assert_eq!(out, &["1", "9"]);
}

#[test]
fn test_min_by_and_max_by_extract_key() {
    let out = run_prints(
        r#"
        fun main() {
            data class Item(val id: Int, val label: String)
            val values = listOf(Item(3, "z"), Item(1, "x"), Item(2, "y"))
            println(values.minByOrNull { it.id }?.label)
            println(values.maxByOrNull { it.id }?.label)
        }
    "#,
    );
    assert_eq!(out, &["x", "z"]);
}

#[test]
fn test_binary_search_existing_sorted() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(1, 3, 5, 7)
            println(values.binarySearch(5))
            println(values.binarySearch(6))
        }
    "#,
    );
    assert_eq!(out, &["2", "-4"]);
}

#[test]
fn test_binary_search_range() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(1, 3, 5, 7, 9)
            println(values.binarySearch(5, 0, 2))
            println(values.binarySearch(3, 0, 2))
        }
    "#,
    );
    assert_eq!(out, &["-3", "1"]);
}

#[test]
fn test_partition_by_parity() {
    let out = run_prints(
        r#"
        fun main() {
            val (even, odd) = listOf(1, 2, 3, 4, 5).partition { it % 2 == 0 }
            println(even.joinToString(","))
            println(odd.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,4", "1,3,5"]);
}

#[test]
fn test_distinct_keeps_first_occurrence_order() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(1, 2, 2, 3, 1, 4, 3)
            println(values.distinct().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3,4"]);
}

#[test]
fn test_distinct_by_first_char() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf("apple", "ant", "banana", "car")
            println(values.distinctBy { it[0] }.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["apple,banana,car"]);
}

#[test]
fn test_sorted_set_projection_to_list() {
    let out = run_prints(
        r#"
        fun main() {
            val values = sortedSetOf("b", "a", "c")
            println(values.toList().joinToString(","))
            println(values.toMutableList()[0])
        }
    "#,
    );
    assert_eq!(out, &["a,b,c", "a"]);
}

#[test]
fn test_sorted_set_with_comparator() {
    let out = run_prints(
        r#"
        fun main() {
            val values = sortedSetOf(compareBy<String> { it.length }.thenBy { it }, "bbb", "cc", "a", "ddd")
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["a,cc,bbb,ddd"]);
}

#[test]
fn test_fold_left_sum_with_sorted_input() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(9, 1, 5, 3).sorted()
            val total = values.fold(0) { acc, n -> acc + n }
            println(values.joinToString(","))
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["1,3,5,9", "18"]);
}

#[test]
fn test_fold_right_concatenate() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf("a", "b", "c").sorted()
            val out = values.foldRight("") { value, acc -> value + acc }
            println(out)
        }
    "#,
    );
    assert_eq!(out, &["abc"]);
}

#[test]
fn test_reduce_right_product_associative() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(2, 3, 4).sortedDescending()
            val product = values.reduce { acc, value -> acc * value }
            println(values.joinToString(","))
            println(product)
        }
    "#,
    );
    assert_eq!(out, &["4,3,2", "24"]);
}

#[test]
fn test_reduce_with_indexed_variant() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(4, 5, 6).sorted()
            val out = values.foldIndexed("") { index, acc, value ->
                if (index == 0) value.toString() else "${'$'}{acc}-${'$'}value"
            }
            println(out)
        }
    "#,
    );
    assert_eq!(out, &["4-5-6"]);
}

#[test]
fn test_scan_accumulation_order() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(1, 2, 3).sorted()
            val accumulated = values.runningReduce { acc, n -> acc + n }
            println(accumulated.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,3,6"]);
}

#[test]
fn test_sorted_drop_take_consistency() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(9, 4, 7, 1).sorted()
            println(values.drop(2).joinToString(","))
            println(values.take(2).joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["7,9", "1,4"]);
}

#[test]
fn test_sorted_slice_windowed() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(10, 3, 5, 7).sorted()
            val windows = values.windowed(2)
            println(windows.size)
            println(windows[0].joinToString(","))
            println(windows[1].joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["3", "1,3", "3,5"]);
}

#[test]
fn test_sorted_chunked_projection() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf(5, 2, 9, 1, 8, 3).sorted()
            val chunked = values.chunked(3).joinToString("|") { it.joinToString(",") }
            println(chunked)
        }
    "#,
    );
    assert_eq!(out, &["1,2,3|5,8,9"]);
}

#[test]
fn test_map_indexed_sorting_preserves_indexed_payload() {
    let out = run_prints(
        r#"
        fun main() {
            val values = listOf("c", "a", "b").withIndex().toList().sortedBy { it.value }
            println(values.map { "${'$'}{it.index}:${'$'}{it.value}" }.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1:a,2:b,0:c"]);
}
