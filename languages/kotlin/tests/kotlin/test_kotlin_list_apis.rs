kotlin_run_cases! {
    test_list_basic_lookup => (r##"
        fun main() {
            val list = listOf(4, 5, 6)
            println(list[0])
            println(list.size)
            println(list[2])
        }
    "##, &[
        "4",
        "3",
        "6",
    ]),
    test_list_first_last => (r##"
        fun main() {
            val list = listOf("a", "b", "c")
            println(list.first())
            println(list.last())
        }
    "##, &[
        "a",
        "c",
    ]),
    test_list_first_or_null => (r##"
        fun main() {
            val list = listOf(1)
            println(list.firstOrNull())
            println(emptyList<Int>().firstOrNull() ?: "none")
        }
    "##, &[
        "1",
        "none",
    ]),
    test_list_last_or_null => (r##"
        fun main() {
            val list = listOf(1, 2, 3)
            println(list.lastOrNull())
            println(emptyList<Int>().lastOrNull() ?: "none")
        }
    "##, &[
        "3",
        "none",
    ]),
    test_list_element_at => (r##"
        fun main() {
            val list = listOf(8, 9, 10)
            println(list.elementAt(1))
            println(list.elementAtOrNull(9) ?: "na")
        }
    "##, &[
        "9",
        "na",
    ]),
    test_list_slice_simple => (r##"
        fun main() {
            val list = listOf(1, 2, 3, 4, 5)
            val slice = list.slice(1..3)
            println(slice.joinToString(","))
        }
    "##, &[
        "2,3,4",
    ]),
    test_list_slice_int_range => (r##"
        fun main() {
            val list = listOf("a", "b", "c", "d")
            val slice = list.slice(IntRange(0, 2))
            println(slice.joinToString(""))
        }
    "##, &[
        "abc",
    ]),
    test_list_sublist => (r##"
        fun main() {
            val list = listOf(1, 2, 3, 4)
            val sub = list.subList(1, 3)
            println(sub.joinToString(","))
        }
    "##, &[
        "2,3",
    ]),
    test_list_take_drop => (r##"
        fun main() {
            val list = listOf(1, 2, 3, 4, 5)
            println(list.take(3).joinToString(","))
            println(list.drop(3).joinToString(","))
        }
    "##, &[
        "1,2,3",
        "4,5",
    ]),
    test_list_take_last_drop_last => (r##"
        fun main() {
            val list = listOf(1, 2, 3, 4)
            println(list.takeLast(2).joinToString(","))
            println(list.dropLast(2).joinToString(","))
        }
    "##, &[
        "3,4",
        "1,2",
    ]),
    test_list_reversed => (r##"
        fun main() {
            val list = listOf("a", "b", "c")
            println(list.reversed().joinToString(""))
        }
    "##, &[
        "cba",
    ]),
    test_list_sorted_descending => (r##"
        fun main() {
            val list = listOf(3, 1, 4, 2)
            println(list.sortedDescending().joinToString(","))
        }
    "##, &[
        "4,3,2,1",
    ]),
    test_list_as_reversed => (r##"
        fun main() {
            val list = listOf(7, 8, 9)
            println(list.asReversed().joinToString(","))
            println(list.joinToString(","))
        }
    "##, &[
        "9,8,7",
        "7,8,9",
    ]),
    test_list_contains => (r##"
        fun main() {
            val list = listOf("x", "y", "z")
            println(list.contains("y"))
            println(list.contains("q"))
        }
    "##, &[
        "true",
        "false",
    ]),
    test_list_contains_all => (r##"
        fun main() {
            val list = listOf(1, 2, 3, 4)
            println(list.containsAll(listOf(1, 3)))
            println(list.containsAll(listOf(1, 7)))
        }
    "##, &[
        "true",
        "false",
    ]),
    test_list_index_of => (r##"
        fun main() {
            val list = listOf(10, 20, 30, 20)
            println(list.indexOf(20))
            println(list.lastIndexOf(20))
            println(list.indexOf(99))
        }
    "##, &[
        "1",
        "3",
        "-1",
    ]),
    test_list_count_matches => (r##"
        fun main() {
            val list = listOf(1, 2, 3, 4, 5)
            println(list.count { it % 2 == 0 })
            println(list.count { it > 4 })
        }
    "##, &[
        "2",
        "1",
    ]),
    test_list_any_all_none => (r##"
        fun main() {
            val list = listOf(1, 2, 3)
            println(list.any { it > 2 })
            println(list.all { it > 0 })
            println(list.none { it < 0 })
        }
    "##, &[
        "true",
        "true",
        "true",
    ]),
    test_list_find_first => (r##"
        fun main() {
            val list = listOf(5, 6, 7)
            println(list.find { it > 5 })
            println(list.findLast { it < 6 } ?: "none")
        }
    "##, &[
        "6",
        "5",
    ]),
    test_list_reduce_sum => (r##"
        fun main() {
            val list = listOf(1, 2, 3, 4)
            println(list.reduce { acc, value -> acc + value })
        }
    "##, &[
        "10",
    ]),
    test_list_fold_seed => (r##"
        fun main() {
            val list = listOf(1, 2, 3)
            println(list.fold("start") { acc, value -> acc + value.toString() })
        }
    "##, &[
        "start123",
    ]),
    test_list_sum_max_min => (r##"
        fun main() {
            val list = listOf(3, 9, 1, 5)
            println(list.sum())
            println(list.maxOrNull())
            println(list.minOrNull())
        }
    "##, &[
        "18",
        "9",
        "1",
    ]),
    test_list_join_to_string_custom => (r##"
        fun main() {
            val list = listOf(1, 2, 3)
            println(list.joinToString(prefix = "[", postfix = "]", separator = ":"))
        }
    "##, &[
        "[1:2:3]",
    ]),
    test_list_windowed => (r##"
        fun main() {
            val list = listOf(1, 2, 3, 4)
            println(list.windowed(2).joinToString("|") { it.joinToString(",") })
        }
    "##, &[
        "1,2|2,3|3,4",
    ]),
    test_list_zip_with => (r##"
        fun main() {
            val a = listOf("a", "b", "c")
            val b = listOf(1, 2, 3, 4)
            println(a.zip(b).joinToString(",") { it.first + it.second.toString() })
        }
    "##, &[
        "a1,b2,c3",
    ]),
    test_list_flat_map => (r##"
        fun main() {
            val list = listOf(listOf(1, 2), listOf(3, 4))
            val out = list.flatMap { it.map { n -> n * 2 } }
            println(out.joinToString(","))
        }
    "##, &[
        "2,4,6,8",
    ]),
    test_list_map_not_null => (r##"
        fun main() {
            val list = listOf(1, null, 2, null, 3)
            val out = list.mapNotNull { it?.toString() }.joinToString(",")
            println(out)
        }
    "##, &[
        "1,2,3",
    ]),
    test_list_distinct => (r##"
        fun main() {
            val list = listOf(1, 1, 2, 2, 3)
            println(list.distinct().joinToString(","))
        }
    "##, &[
        "1,2,3",
    ]),
    test_list_distinct_by => (r##"
        fun main() {
            val list = listOf("aa", "ab", "bc", "bd", "c")
            val out = list.distinctBy { it.length }
            println(out.joinToString(","))
        }
    "##, &[
        "aa,bc,c",
    ]),
    test_list_retain_if_even => (r##"
        fun main() {
            val list = mutableListOf(1, 2, 3, 4, 5)
            list.retainAll { it % 2 == 0 }
            println(list.joinToString(","))
        }
    "##, &[
        "2,4",
    ]),
    test_list_remove_if => (r##"
        fun main() {
            val list = mutableListOf(1, 2, 3, 4, 5)
            list.removeIf { it < 3 }
            println(list.joinToString(","))
        }
    "##, &[
        "3,4,5",
    ]),
    test_list_replace_all => (r##"
        fun main() {
            val list = mutableListOf(1, 2, 3)
            list.replaceAll { it * 3 }
            println(list.joinToString(","))
        }
    "##, &[
        "3,6,9",
    ]),
    test_list_mutable_add_remove => (r##"
        fun main() {
            val list = mutableListOf(1, 2)
            list.add(3)
            list.add(1, 9)
            list.removeAt(0)
            println(list.joinToString(","))
        }
    "##, &[
        "9,2,3",
    ]) }
