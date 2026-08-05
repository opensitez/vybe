kotlin_run_cases! {
    test_chunked_two => (r##"
        fun main() {
            val values = (1..7).toList()
            val chunks = values.chunked(3)
            println(chunks.size)
            println(chunks[0].joinToString(","))
            println(chunks.last().joinToString(","))
        }
    "##, vec![String::from("3"), String::from("1,2,3"), String::from("7")]),
    test_chunked_transform => (r##"
        fun main() {
            val values = listOf(1, 2, 3, 4, 5)
            val transformed = values.chunked(2) { it.sum() }
            println(transformed.joinToString(","))
        }
    "##, vec![String::from("3,7,5")]),
    test_windowed_size => (r##"
        fun main() {
            val values = listOf(1, 2, 3, 4, 5)
            val windows = values.windowed(3)
            println(windows.size)
            println(windows[0].joinToString(","))
            println(windows[1].joinToString(","))
        }
    "##, vec![String::from("3"), String::from("1,2,3"), String::from("2,3,4")]),
    test_windowed_partial => (r##"
        fun main() {
            val values = listOf(1, 2, 3)
            val windows = values.windowed(2, partialWindows = true)
            println(windows.size)
            println(windows[2].joinToString(","))
        }
    "##, vec![String::from("3"), String::from("3")]),
    // ^ partialWindows keeps the tails: [1,2], [2,3], [3] — size 3, and
    // windows[2] is the partial [3] (real Kotlin agrees).
    test_windowed_step => (r##"
        fun main() {
            val values = listOf(1, 2, 3, 4, 5, 6)
            val windows = values.windowed(2, step = 2)
            println(windows.size)
            println(windows.joinToString("|") { it.joinToString(",") })
        }
    "##, vec![String::from("3"), String::from("1,2|3,4|5,6")]),
    test_zip_two_lists => (r##"
        fun main() {
            val a = listOf(1, 2, 3)
            val b = listOf("a", "b", "c")
            val zipped = a.zip(b)
            println(zipped.size)
            println(zipped[0].first.toString())
            println(zipped[2].second)
        }
    "##, vec![String::from("3"), String::from("1"), String::from("c")]),
    test_zip_with => (r##"
        fun main() {
            val nums = listOf(1, 2, 3)
            val chars = listOf("a", "b", "c")
            val out = nums.zip(chars) { n, c -> "$c$n" }
            println(out.joinToString(","))
        }
    "##, vec![String::from("a1,b2,c3")]),
    test_unzip => (r##"
        fun main() {
            val pairs = listOf(1 to "a", 2 to "b")
            val (numbers, letters) = pairs.unzip()
            println(numbers.joinToString(","))
            println(letters.joinToString(","))
        }
    "##, vec![String::from("1,2"), String::from("a,b")]),
    test_flatten_nested => (r##"
        fun main() {
            val nested = listOf(listOf(1, 2), listOf(3), listOf(4, 5))
            println(nested.flatten().joinToString(","))
        }
    "##, vec![String::from("1,2,3,4,5")]),
    test_flat_map_lists => (r##"
        fun main() {
            val source = listOf(listOf(1, 2), listOf(3, 4))
            val mapped = source.flatMap { it.map { v -> v * 10 } }
            println(mapped.joinToString(","))
        }
    "##, vec![String::from("10,20,30,40")]),
    test_distinct_and_retain => (r##"
        fun main() {
            val source = listOf(1, 1, 2, 3, 3, 4)
            println(source.distinct().joinToString(","))
            println(source.distinctBy { it % 2 }.joinToString(","))
        }
    "##, vec![String::from("1,2,3,4"), String::from("1,2")]),
    test_group_by => (r##"
        fun main() {
            val nums = listOf(1, 2, 3, 4)
            val grouped = nums.groupBy { it % 2 == 0 }
            val evens = grouped[true] ?: emptyList<Int>()
            val odds = grouped[false] ?: emptyList<Int>()
            println(evens.joinToString(","))
            println(odds.joinToString(","))
        }
    "##, vec![String::from("2,4"), String::from("1,3")]),
    test_map_not_null => (r##"
        fun main() {
            val values = listOf("1", null, "x", "3")
            val out = values.mapNotNull { it?.toIntOrNull() }
            println(out.joinToString(","))
        }
    "##, vec![String::from("1,3")]),
    test_associate_from_chunked => (r##"
        fun main() {
            val pairs = listOf(1, 2, 3)
            val map = pairs.chunked(1).associate { chunk ->
                val key = chunk[0]
                key to "v$key"
            }
            println(map.size)
            println(map[2])
        }
    "##, vec![String::from("3"), String::from("v2")]),
    test_windowed_join => (r##"
        fun main() {
            val values = "abccba"
            val windows = values.windowed(2, partialWindows = false)
            println(windows.size)
            println(windows.joinToString("|"))
        }
    "##, vec![String::from("5"), String::from("ab|bc|cc|cb|ba")]),
}
