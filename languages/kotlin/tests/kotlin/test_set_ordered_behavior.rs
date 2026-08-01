use crate::helpers::run_prints;

#[test]
fn test_linked_hash_set_keeps_insertion_and_removal_order() {
    let out = run_prints(
        r#"
        fun main() {
            val values = linkedSetOf(3, 1, 2, 4)
            values.remove(1)
            values.add(1)
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["3,2,4,1"]);
}

#[test]
fn test_linked_set_reinsert_existing_is_noop_for_order() {
    let out = run_prints(
        r#"
        fun main() {
            val values = linkedSetOf(1, 2, 3)
            values.add(2)
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3"]);
}

#[test]
fn test_tree_set_natural_ordering() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.TreeSet<Int>()
            values.add(5)
            values.add(1)
            values.add(3)
            println(values.joinToString(","))
            println(values.first())
            println(values.last())
        }
    "#,
    );
    assert_eq!(out, &["1,3,5", "1", "5"]);
}

#[test]
fn test_tree_set_descending_set_view() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.TreeSet<Int>()
            values.addAll(listOf(4, 1, 3, 2))
            val down = values.descendingSet()
            println(down.joinToString(","))
            println(down.first())
            println(down.last())
        }
    "#,
    );
    assert_eq!(out, &["4,3,2,1", "4", "1"]);
}

#[test]
fn test_sorted_set_with_custom_comparator() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.TreeSet<String>(compareByDescending { it })
            values.add("a")
            values.add("c")
            values.add("b")
            println(values.joinToString(","))
            println(values.first())
            println(values.last())
        }
    "#,
    );
    assert_eq!(out, &["c,b,a", "c", "a"]);
}

#[test]
fn test_sorted_set_floor_and_ceiling() {
    let out = run_prints(
        r#"
        fun main() {
            val values = java.util.TreeSet(listOf(2, 4, 6, 8))
            println(values.floor(5))
            println(values.ceiling(5))
            println(values.lower(6))
            println(values.higher(6))
        }
    "#,
    );
    assert_eq!(out, &["4", "6", "4", "8"]);
}

#[test]
fn test_linked_hash_set_to_typed_array() {
    let out = run_prints(
        r#"
        fun main() {
            val values = linkedSetOf(7, 1, 9)
            val arr = values.toTypedArray()
            println(arr.joinToString(","))
            val round = arr.toList().toMutableSet()
            round.add(4)
            println(round.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["7,1,9", "7,1,9,4"]);
}

#[test]
fn test_set_intersection_preserves_left_order_when_iterated() {
    let out = run_prints(
        r#"
        fun main() {
            val left = linkedSetOf(1, 2, 3, 4)
            val right = linkedSetOf(4, 2)
            val inter = left.intersect(right)
            println(inter.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,4"]);
}

#[test]
fn test_set_union_projection_order() {
    let out = run_prints(
        r#"
        fun main() {
            val a = linkedSetOf(4, 1)
            val b = linkedSetOf(2, 3)
            val all = a union b
            println(all.joinToString(","))
            println(all.size)
        }
    "#,
    );
    assert_eq!(out, &["4,1,2,3", "4"]);
}

#[test]
fn test_set_subtract_and_distinct() {
    let out = run_prints(
        r#"
        fun main() {
            val a = linkedSetOf(1, 2, 3, 4)
            val b = linkedSetOf(2, 4, 6)
            println((a - b).joinToString(","))
            val dup = listOf(1,1,2,2,3)
            println(dup.toSet().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,3", "1,2,3"]);
}

#[test]
fn test_linked_set_mutation_during_iteration() {
    let out = run_prints(
        r#"
        fun main() {
            val values = linkedSetOf(1, 2, 3)
            val outValues = StringBuilder()
            val it = values.iterator()
            while (it.hasNext()) {
                val n = it.next()
                if (n == 2) {
                    it.remove()
                }
                outValues.append(n)
            }
            println(outValues.toString())
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["123", "1,3"]);
}
