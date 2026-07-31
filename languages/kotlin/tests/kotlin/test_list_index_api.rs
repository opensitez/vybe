use crate::helpers::run_prints;

#[test]
fn test_list_size_and_last_index()
{
    let out = run_prints(r#"
        fun main() {
            val values = listOf(10, 20, 30, 40)
            println(values.size)
            println(values.lastIndex)
        }
    "#);
    assert_eq!(out, &["4", "3"]);
}

#[test]
fn test_list_indexing_by_operator() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableListOf("a", "b", "c")
            println(values[0])
            println(values[1])
            println(values[2])
        }
    "#);
    assert_eq!(out, &["a", "b", "c"]);
}

#[test]
fn test_list_update_through_index_operator() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableListOf(1, 2, 3)
            values[0] = 7
            values[2] = 9
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["7,2,9"]);
}

#[test]
fn test_list_contains_and_indices_check() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf("x", "y", "z")
            println(values.contains("y"))
            println(values.indices.first())
            println(values.indices.last())
        }
    "#);
    assert_eq!(out, &["true", "0", "2"]);
}

#[test]
fn test_list_first_and_last() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(3, 5, 8)
            println(values.first())
            println(values.last())
            println(values.firstOrNull())
            println(values.lastOrNull())
        }
    "#);
    assert_eq!(out, &["3", "8", "3", "8"]);
}

#[test]
fn test_list_empty_behaviors() {
    let out = run_prints(r#"
        fun main() {
            val values = emptyList<Int>()
            println(values.isEmpty())
            println(values.firstOrNull() ?: "none")
            println(values.lastOrNull() ?: "none")
            println(values.elementAtOrElse(0) { it + 4 })
            println(values.elementAtOrNull(0) ?: "none")
        }
    "#);
    assert_eq!(out, &["true", "none", "none", "4", "none"]);
}

#[test]
fn test_list_sublist_slices_view() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableListOf(1, 2, 3, 4, 5)
            val part = values.subList(1, 4)
            part[0] = 9
            println(values.joinToString(","))
            println(part.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,9,3,4,5", "9,3,4"]);
}

#[test]
fn test_list_slice_by_indices() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(7, 8, 9, 10, 11)
            val part = values.slice(1..3)
            println(part.joinToString(","))
            println(part.size)
        }
    "#);
    assert_eq!(out, &["8,9,10", "3"]);
}

#[test]
fn test_list_index_of_predicate_search() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(1, 2, 3, 2, 4)
            println(values.indexOf(2))
            println(values.lastIndexOf(2))
            println(values.indexOfFirst { it > 2 })
            println(values.indexOfLast { it % 2 == 0 })
        }
    "#);
    assert_eq!(out, &["1", "3", "2", "3"]);
}

#[test]
fn test_list_get_or_else_and_or_null() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(4, 5, 6)
            println(values.getOrElse(1) { 0 })
            println(values.getOrElse(4) { 99 })
            println(values.getOrNull(0) ?: -1)
            println(values.getOrNull(4) ?: -1)
        }
    "#);
    assert_eq!(out, &["5", "99", "4", "-1"]);
}

#[test]
fn test_list_element_access_from_reversed_view() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(1, 2, 3)
            val rev = values.asReversed()
            println(rev[0])
            println(rev[1])
            println(rev[2])
        }
    "#);
    assert_eq!(out, &["3", "2", "1"]);
}

#[test]
fn test_list_iterator_navigation_contract() {
    let out = run_prints(r#"
        fun main() {
            val it = listOf(1, 2, 3).listIterator()
            println(it.hasNext())
            println(it.next())
            println(it.next())
            println(it.hasPrevious())
            println(it.previous())
            println(it.previousIndex())
        }
    "#);
    assert_eq!(out, &["true", "1", "2", "true", "1", "0"]);
}

#[test]
fn test_list_to_typed_array_round_trip() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(4, 5, 6)
            val arr = values.toTypedArray()
            val back = arr.toList()
            println(back.size)
            println(back.joinToString(","))
            println((back === values).toString())
        }
    "#);
    assert_eq!(out, &["3", "4,5,6", "false"]);
}

#[test]
fn test_list_drop_take_boundaries() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(1, 2, 3, 4, 5)
            println(values.take(2).joinToString(","))
            println(values.drop(2).joinToString(","))
            println(values.takeLast(2).joinToString(","))
            println(values.dropLast(4).joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,2", "3,4,5", "4,5", "1"]);
}

#[test]
fn test_list_component_copy_and_mutation_independence() {
    let out = run_prints(r#"
        fun main() {
            val source = mutableListOf(1, 2, 3)
            val copy = source.toMutableList()
            copy.add(4)
            println(source.size)
            println(copy.size)
            println(source.joinToString(","))
            println(copy.joinToString(","))
        }
    "#);
    assert_eq!(out, &["3", "4", "1,2,3", "1,2,3,4"]);
}

#[test]
fn test_list_mutable_iterator_stepwise() {
    let out = run_prints(r#"
        fun main() {
            val iterator = mutableListOf(5, 6, 7, 8).iterator()
            var total = 0
            while (iterator.hasNext()) {
                total += iterator.next()
            }
            println(total)
            println(iterator.hasNext())
        }
    "#);
    assert_eq!(out, &["26", "false"]);
}

#[test]
fn test_list_last_index_after_mutation() {
    let out = run_prints(r#"
        fun main() {
            val values = mutableListOf(1, 2)
            values.add(3)
            println(values.lastIndex)
            values.removeAt(2)
            println(values.lastIndex)
        }
    "#);
    assert_eq!(out, &["2", "1"]);
}

#[test]
fn test_list_slice_empty_range_returns_empty() {
    let out = run_prints(r#"
        fun main() {
            val values = listOf(1, 2, 3, 4)
            val part = values.slice(2 until 2)
            println(part.isEmpty())
            println(part.size)
        }
    "#);
    assert_eq!(out, &["true", "0"]);
}
