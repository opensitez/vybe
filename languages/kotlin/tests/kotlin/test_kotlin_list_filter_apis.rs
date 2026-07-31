kotlin_run_cases! {
    test_filter_predicates => (r##"
        fun main() {
            val nums = listOf(1, 2, 3, 4, 5)
            println(nums.filter { it % 2 == 0 }.joinToString(","))
            println(nums.filterNot { it > 3 }.joinToString(","))
            println(nums.filterIndexed { index, _ -> index % 2 == 0 }.joinToString(","))
        }
    "##, &[
        "2,4",
        "1,2,3",
        "1,3,5",
    ]),
    test_filter_not_null => (r##"
        fun main() {
            val values = listOf("a", null, "b", null, "c")
            println(values.filterNotNull().joinToString(","))
            println(values.count { it == null })
        }
    "##, &[
        "a,b,c",
        "2",
    ]),
    test_filter_with_index_and_range => (r##"
        fun main() {
            val nums = listOf(10, 11, 12, 13)
            val filtered = nums.filterIndexed { idx, value -> idx + value == 11 || idx + value == 15 }
            println(filtered.joinToString(","))
        }
    "##, &[
        "10,12",
    ]),
    test_filter_is_instance => (r##"
        fun main() {
            val values: List<Any> = listOf(1, "a", 2, "b", 3.0)
            val ints = values.filterIsInstance<Int>()
            val strings = values.filterIsInstance<String>()
            println(ints.joinToString(","))
            println(strings.joinToString(","))
        }
    "##, &[
        "1,2",
        "a,b",
    ]),
    test_take_drop_operations => (r##"
        fun main() {
            val nums = listOf(1, 2, 3, 4, 5, 6)
            println(nums.take(3).joinToString(","))
            println(nums.drop(3).joinToString(","))
            println(nums.takeLast(2).joinToString(","))
            println(nums.dropLast(2).joinToString(","))
        }
    "##, &[
        "1,2,3",
        "4,5,6",
        "5,6",
        "1,2,3,4",
    ]),
    test_slice_window_chunk => (r##"
        fun main() {
            val nums = listOf(1, 2, 3, 4, 5)
            println(nums.slice(1..3).joinToString(","))
            println(nums.subList(0, 2).joinToString(","))
            println(nums.sliceArray(2 until 4).joinToString(","))
        }
    "##, &[
        "2,3,4",
        "1,2",
        "3,4",
    ]),
    test_distinct_and_retain => (r##"
        fun main() {
            val nums = listOf(1, 2, 2, 3, 3, 4)
            println(nums.distinct().joinToString(","))
            val words = listOf("a", "ab", "ac", "b")
            println(words.distinctBy { it[0] }.joinToString(","))
        }
    "##, &[
        "1,2,3,4",
        "a,b",
    ]),
    test_retain_and_intersections => (r##"
        fun main() {
            val source = listOf(1, 2, 3, 4)
            val evens = source.filter { it % 2 == 0 }
            val odds = source.filter { it % 2 == 1 }
            println(evens.joinToString(","))
            println(odds.joinToString(","))
            println(source.intersect(evens.toSet()).joinToString(","))
        }
    "##, &[
        "2,4",
        "1,3",
        "2,4",
    ]),
    test_partition_by_condition => (r##"
        fun main() {
            val nums = listOf(5, 10, 11, 20)
            val (evens, odds) = nums.partition { it % 2 == 0 }
            println(evens.joinToString(","))
            println(odds.joinToString(","))
            println(evens.count())
        }
    "##, &[
        "10,20",
        "5,11",
        "2",
    ]),
    test_take_while_drop_while => (r##"
        fun main() {
            val nums = listOf(1, 2, 3, 2, 1)
            println(nums.takeWhile { it < 3 }.joinToString(","))
            println(nums.dropWhile { it < 3 }.joinToString(","))
            println(nums.takeWhile { it > 0 }.joinToString(","))
        }
    "##, &[
        "1,2",
        "3,2,1",
        "1,2,3,2,1",
    ]),
    test_drop_while_empty => (r##"
        fun main() {
            val nums = listOf(1, 2, 3)
            println(nums.dropWhile { it < 0 }.joinToString(","))
            println(nums.takeWhile { false }.joinToString(","))
            println(nums.takeWhile { true }.size)
        }
    "##, &[
        "1,2,3",
        "",
        "3",
    ]),
    test_filter_and_size_shortcuts => (r##"
        fun main() {
            val nums = listOf(1, 2, 3, 4, 5)
            println(nums.all { it > 0 })
            println(nums.any { it > 4 })
            println(nums.none { it > 10 })
        }
    "##, &[
        "true",
        "true",
        "true",
    ]),
}
