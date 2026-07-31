kotlin_run_cases! {
    test_set_of_unique_values => (r##"
        fun main() {
            val set = setOf(1, 2, 2, 3)
            println(set.size)
            println(set.contains(2))
            println(set.contains(9))
        }
    "##, &[
        "3",
        "true",
        "false",
    ]),
    test_set_linked_insertion_order => (r##"
        fun main() {
            val set = linkedSetOf("z", "a", "m", "a")
            println(set.joinToString(","))
            println(set.size)
        }
    "##, &[
        "z,a,m",
        "3",
    ]),
    test_set_sorted_set_order => (r##"
        fun main() {
            val set = sortedSetOf(5, 1, 4, 2, 3)
            println(set.joinToString(","))
            println(set.first())
            println(set.last())
        }
    "##, &[
        "1,2,3,4,5",
        "1",
        "5",
    ]),
    test_mutable_set_add_duplicate => (r##"
        fun main() {
            val set = mutableSetOf("x")
            val first = set.add("x")
            val second = set.add("y")
            println(first)
            println(second)
            println(set.size)
        }
    "##, &[
        "false",
        "true",
        "2",
    ]),
    test_mutable_set_remove_present => (r##"
        fun main() {
            val set = mutableSetOf(1, 2, 3)
            val removed = set.remove(2)
            println(removed)
            println(set.size)
            println(set.contains(2))
        }
    "##, &[
        "true",
        "2",
        "false",
    ]),
    test_mutable_set_remove_absent => (r##"
        fun main() {
            val set = mutableSetOf(1, 2)
            val removed = set.remove(7)
            println(removed)
            println(set.size)
        }
    "##, &[
        "false",
        "2",
    ]),
    test_set_clear => (r##"
        fun main() {
            val set = mutableSetOf(1, 2, 3)
            set.clear()
            println(set.isEmpty())
            println(set.size)
        }
    "##, &[
        "true",
        "0",
    ]),
    test_set_plus_operator => (r##"
        fun main() {
            val a = setOf(1, 2, 3)
            val b = setOf(3, 4)
            val c = a + b
            println(c.size)
            println(c.joinToString(","))
        }
    "##, &[
        "4",
        "1,2,3,4",
    ]),
    test_set_minus_operator => (r##"
        fun main() {
            val a = linkedSetOf(1, 2, 3, 4)
            val b = a - listOf(2, 4)
            println(b.size)
            println(b.joinToString(","))
        }
    "##, &[
        "2",
        "1,3",
    ]),
    test_set_union => (r##"
        fun main() {
            val a = setOf(1, 2)
            val b = setOf(2, 3, 4)
            val c = a.union(b)
            println(c.size)
            println(c.contains(4))
        }
    "##, &[
        "4",
        "true",
    ]),
    test_set_intersection => (r##"
        fun main() {
            val a = setOf(1, 2, 3, 4)
            val b = setOf(3, 4, 5)
            val c = a.intersect(b)
            println(c.size)
            println(c.joinToString(","))
        }
    "##, &[
        "2",
        "3,4",
    ]),
    test_set_subtract => (r##"
        fun main() {
            val a = setOf(1, 2, 3, 4)
            val b = setOf(2, 4)
            val c = a - b
            println(c.joinToString(","))
        }
    "##, &[
        "1,3",
    ]),
    test_set_retain_all => (r##"
        fun main() {
            val set = mutableSetOf(1, 2, 3, 4)
            val changed = set.retainAll(listOf(2, 4))
            println(changed)
            println(set.joinToString(","))
        }
    "##, &[
        "true",
        "2,4",
    ]),
    test_set_remove_all => (r##"
        fun main() {
            val set = mutableSetOf(1, 2, 3, 4)
            val changed = set.removeAll(listOf(1, 3))
            println(changed)
            println(set.size)
            println(set.contains(3))
        }
    "##, &[
        "true",
        "2",
        "false",
    ]),
    test_set_add_all => (r##"
        fun main() {
            val set = mutableSetOf(1, 2)
            val changed = set.addAll(listOf(2, 3, 4))
            println(changed)
            println(set.joinToString(","))
        }
    "##, &[
        "true",
        "1,2,3,4",
    ]),
    test_set_contains_all => (r##"
        fun main() {
            val set = setOf(1, 2, 3)
            println(set.containsAll(listOf(1, 3)))
            println(set.containsAll(listOf(1, 4)))
        }
    "##, &[
        "true",
        "false",
    ]),
    test_set_any_none_all => (r##"
        fun main() {
            val set = setOf(1, 2, 4)
            println(set.any { it > 3 })
            println(set.none { it < 0 })
            println(set.all { it < 10 })
        }
    "##, &[
        "true",
        "true",
        "true",
    ]),
    test_set_singleton => (r##"
        fun main() {
            val set = setOf(7)
            println(set.size)
            println(set.first())
            println(set.last())
        }
    "##, &[
        "1",
        "7",
        "7",
    ]),
    test_set_first_and_last_in_sorted => (r##"
        fun main() {
            val set = sortedSetOf(9, 1, 4)
            println(set.first())
            println(set.last())
        }
    "##, &[
        "1",
        "9",
    ]),
    test_set_distinct_preserved_after_map => (r##"
        fun main() {
            val set = setOf(1, 2, 3)
            val doubled = set.map { it * 2 }.toSet()
            println(doubled.joinToString(","))
        }
    "##, &[
        "2,4,6",
    ]),
    test_set_of_not_in => (r##"
        fun main() {
            val set = setOf(1, 2, 3)
            println(4 !in set)
            println(2 !in set)
        }
    "##, &[
        "true",
        "false",
    ]),
    test_set_to_list_roundtrip => (r##"
        fun main() {
            val set = linkedSetOf("a", "b", "c")
            val list = set.toList()
            println(list.joinToString(","))
            println(set.joinToString(","))
        }
    "##, &[
        "a,b,c",
        "a,b,c",
    ]),
    test_set_hash_code_stable_order => (r##"
        fun main() {
            val set = setOf("x", "y", "z")
            println(set.size)
            println(set.contains("y"))
        }
    "##, &[
        "3",
        "true",
    ]),
    test_set_max_min => (r##"
        fun main() {
            val set = sortedSetOf(10, 5, 12, 7)
            println(set.minOrNull())
            println(set.maxOrNull())
        }
    "##, &[
        "5",
        "12",
    ]),
    test_set_fold_sum => (r##"
        fun main() {
            val set = setOf(1, 2, 3)
            println(set.fold(0) { acc, v -> acc + v })
        }
    "##, &[
        "6",
    ]),
    test_set_sum_reduce => (r##"
        fun main() {
            val set = setOf(2, 4, 6)
            println(set.sum())
            println(set.reduce { a, b -> a + b })
        }
    "##, &[
        "12",
        "12",
    ]),
    test_set_map_is_empty_after_remove_last => (r##"
        fun main() {
            val set = mutableSetOf(1)
            set.remove(1)
            println(set.isEmpty())
        }
    "##, &[
        "true",
    ]),
    test_set_empty_intersection => (r##"
        fun main() {
            val set = emptySet<Int>().intersect(setOf(1, 2))
            println(set.isEmpty())
            println(set.size)
        }
    "##, &[
        "true",
        "0",
    ]),
    test_set_empty_union => (r##"
        fun main() {
            val set = setOf<Int>().union(setOf(1, 2, 2))
            println(set.size)
            println(set.joinToString(","))
        }
    "##, &[
        "2",
        "1,2",
    ]),
    test_set_partition_like => (r##"
        fun main() {
            val set = setOf(1, 2, 3, 4)
            val out = set.partition { it % 2 == 0 }
            println(out.first.joinToString(","))
            println(out.second.joinToString(","))
        }
    "##, &[
        "2,4",
        "1,3",
    ]),
    test_set_filter_even => (r##"
        fun main() {
            val set = setOf(1, 2, 3, 4, 5)
            val evens = set.filter { it % 2 == 0 }
            println(evens.joinToString(","))
            println(evens.size)
        }
    "##, &[
        "2,4",
        "2",
    ]),
    test_set_join_to_string => (r##"
        fun main() {
            val set = linkedSetOf("s", "t", "u")
            println(set.joinToString("|"))
        }
    "##, &[
        "s|t|u",
    ]),
    test_set_to_mutable_set => (r##"
        fun main() {
            val immutable = setOf(1, 2)
            val mutable = immutable.toMutableSet()
            mutable.add(3)
            println(immutable.size)
            println(mutable.size)
            println(mutable.contains(3))
        }
    "##, &[
        "2",
        "3",
        "true",
    ]),
    test_set_contains_sequence => (r##"
        fun main() {
            val set = setOf(5, 6, 7)
            val all = sequenceOf(5, 8, 7).all { set.contains(it) }
            println(all)
        }
    "##, &[
        "false",
    ]),
}
