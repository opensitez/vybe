kotlin_run_cases! {
    test_grouping_by_length => (r#"
        fun main() {
            val words = listOf("a", "bee", "cat", "deer", "dog")
            val grouped = words.groupBy { it.length }
            val short = grouped[1]!!.joinToString(",")
            val long = grouped[3]!!.joinToString(",")
            println(short)
            println(long)
        }
    "#, &["a", "cat,dog"]),
    test_grouping_by_first_char => (r#"
        fun main() {
            val words = listOf("apple", "apricot", "banana", "blue")
            val grouped = words.groupBy { it.first() }
            println(grouped['a']!!.joinToString(","))
            println(grouped['b']!!.size)
        }
    "#, &["apple,apricot", "2"]),
    test_grouping_by_even_odd => (r#"
        fun main() {
            val values = listOf(1,2,3,4,5,6)
            val grouped = values.groupBy { it % 2 == 0 }
            println(grouped[true]!!.joinToString(","))
            println(grouped[false]!!.joinToString(","))
        }
    "#, &["2,4,6", "1,3,5"]),
    test_group_by_keeping_order => (r#"
        fun main() {
            val values = listOf("bb", "a", "c", "aa", "b")
            val grouped = values.groupBy { it.length }
            println(grouped[1]!!.joinToString("|"))
            println(grouped[2]!!.joinToString("|"))
        }
    "#, &["a|c|b", "bb|aa"]),
    test_grouping_by_nested_selector => (r#"
        fun main() {
            val points = listOf(1, 12, 3, 20, 17)
            val grouped = points.groupBy { if (it >= 10) "high" else "low" }
            println(grouped["high"]!!.size)
            println(grouped["low"]!!.size)
        }
    "#, &["2", "3"]),
    test_grouping_to_existing_map => (r#"
        fun main() {
            val source = listOf("k", "kk", "kotlin")
            val out = linkedMapOf<Int, MutableList<String>>()
            source.groupByTo(out, { it.length }, { it })
            println(out[1]!!.joinToString(","))
            println(out[6]!!.first())
            println(out.size)
        }
    "#, &["k", "kotlin", "3"]),
    test_grouping_with_null_key => (r#"
        fun main() {
            val values = listOf("a", "", "b", "")
            val grouped = values.groupBy { if (it.isEmpty()) null else it }
            println(grouped[null]!!.size)
            println(grouped["a"]!!.size)
        }
    "#, &["2", "1"]),
    test_grouping_counting_empty_input => (r#"
        fun main() {
            val grouped = emptyList<Int>().groupBy { it }
            println(grouped.isEmpty())
        }
    "#, &["true"]),
    test_grouping_empty_strings => (r#"
        fun main() {
            val values = listOf("", "", "x")
            val grouped = values.groupBy { it }
            println(grouped[""]!!.size)
            println(grouped["x"]!!.size)
        }
    "#, &["2", "1"]),
    test_grouping_by_range => (r#"
        fun main() {
            val values = (1..10).toList()
            val grouped = values.groupBy { if (it <= 3) "small" else "big" }
            println(grouped["small"]!!.sum())
            println(grouped["big"]!!.size)
        }
    "#, &["6", "7"]),
    test_grouping_by_modulo_three => (r#"
        fun main() {
            val values = listOf(1, 2, 3, 4, 5, 6)
            val grouped = values.groupBy { it % 3 }
            println(grouped[0]!!.joinToString(","))
            println(grouped[1]!!.joinToString(","))
            println(grouped[2]!!.joinToString(","))
        }
    "#, &["3,6", "1,4", "2,5"]),
    test_grouping_by_string_length_multiple => (r#"
        fun main() {
            val values = listOf("ant", "bear", "cat", "deer", "eel")
            val grouped = values.groupBy(String::length)
            val keys = grouped.keys.toList().sorted()
            println(keys.joinToString(","))
            println(grouped[3]!!.size)
        }
    "#, &["3,4", "2"]),
    test_grouping_by_chars => (r#"
        fun main() {
            val values = listOf("a,b", "c,d", "a:e")
            val grouped = values.groupBy { it[0] }
            println(grouped['a']!!.size)
            println(grouped['c']!!.first())
        }
    "#, &["2", "c,d"]),
    test_grouping_each_count => (r#"
        fun main() {
            val values = listOf("aa", "bb", "a", "bb", "a", "a")
            val counts = values.groupingBy { it.length }.eachCount()
            println(counts[1])
            println(counts[2])
        }
    "#, &["1", "5"]),
    test_grouping_each_count_empty => (r#"
        fun main() {
            val counts = emptyList<String>().groupingBy { it.length }.eachCount()
            println(counts.isEmpty())
        }
    "#, &["true"]),
    test_grouping_by_fold_sum => (r#"
        fun main() {
            val values = listOf(1, 3, 2, 4, 5, 7)
            val sums = values.groupingBy { it % 2 == 0 }.fold(0) { acc, v -> acc + v }
            println(sums[true])
            println(sums[false])
        }
    "#, &["6", "16"]),
    test_grouping_aggregate_chars => (r#"
        fun main() {
            val words = listOf("apple", "ape", "bat", "ball", "cat")
            val out = words.groupingBy { it.first() }
                .aggregate { key, accumulator, element, first ->
                    val value = (accumulator ?: "") + element.length.toString()
                    value
                }
            println(out['a'])
            println(out['b'])
            println(out['c'])
        }
    "#, &["42", "12", "3"]),
    test_grouping_by_index => (r#"
        fun main() {
            val values = listOf("a", "b", "c", "aa")
            val grouped = values.withIndex().groupBy { it.index % 2 }
            println(grouped[0]!!.map { it.value }.joinToString(","))
            println(grouped[1]!!.map { it.value }.joinToString(","))
        }
    "#, &["a,c", "b,aa"]),
    test_grouping_by_index_even_odd_count => (r#"
        fun main() {
            val grouped = (1..7).withIndex().groupBy { it.index % 2 }
            println(grouped[0]!!.size)
            println(grouped[1]!!.size)
        }
    "#, &["4", "3"]),
    test_grouping_map_view_keys => (r#"
        fun main() {
            val words = listOf("x", "yy", "zzz")
            val grouped = words.groupBy { it.length }
            val keys = grouped.keys.sorted()
            println(keys.joinToString(","))
            println(grouped[2]!!.first())
        }
    "#, &["1,2,3", "yy"]),
    test_grouping_by_empty_string_values => (r#"
        fun main() {
            val grouped = listOf("", "a", "", "bb", "", "bbb").groupBy { it.length }
            println(grouped[0]!!.size)
            println(grouped[3]!!.size)
        }
    "#, &["3", "1"]),
    test_grouping_by_unicode_char => (r#"
        fun main() {
            val words = listOf("ß", "ss", "☃")
            val grouped = words.groupBy { it[0] }
            println(grouped['ß']!!.size)
            println(grouped['☃']!!.first())
        }
    "#, &["1", "☃"]),
    test_grouping_by_boolean_key => (r#"
        fun main() {
            val values = listOf(1, 2, 3, 4, 5)
            val grouped = values.groupBy { it > 3 }
            println(grouped[true]!!.joinToString(","))
            println(grouped[false]!!.joinToString(","))
        }
    "#, &["4,5", "1,2,3"]),
    test_grouping_by_custom_object_keys => (r#"
        fun main() {
            val data = listOf("ab", "cd", "efg", "hi")
            val grouped = data.groupBy { Pair(it.length, it[0]) }
            println(grouped[Pair(2, 'a')]!![0])
            println(grouped[Pair(2, 'c')]!![0])
        }
    "#, &["ab", "cd"]),
    test_grouping_by_lexicographic_case => (r#"
        fun main() {
            val values = listOf("a", "A", "b", "B")
            val grouped = values.groupBy { it.lowercase().first() }
            println(grouped['a']!!.size)
            println(grouped['b']!!.size)
        }
    "#, &["2", "2"]),
    test_grouping_by_digit_chars => (r#"
        fun main() {
            val values = listOf("a1", "b2", "c3", "d4")
            val grouped = values.groupBy { it.last() }
            println(grouped['1']!!.first())
            println(grouped['4']!!.first())
        }
    "#, &["a1", "d4"]),
    test_grouping_each_count_ordered_map => (r#"
        fun main() {
            val values = listOf("aa", "bb", "a", "cc", "d")
            val counts = values.groupingBy { it.length }.eachCount().toList().sortedBy { it.first }
            println(counts.joinToString("|") { it.first.toString() + ":" + it.second })
        }
    "#, &["1:2|2:3"]),
    test_grouping_with_mixed_types => (r#"
        fun main() {
            val values = listOf("1", "two", "3", "four")
            val grouped = values.groupBy { if (it.all { ch -> ch.isDigit() }) "digit" else "word" }
            println(grouped["digit"]!!.joinToString(","))
            println(grouped["word"]!!.size)
        }
    "#, &["1,3", "2"]),
    test_grouping_range_map_values => (r#"
        fun main() {
            val values = (0..6).toList()
            val grouped = values.groupBy { it / 2 }
            println(grouped[0]!!.joinToString(","))
            println(grouped[3]!!.joinToString(","))
        }
    "#, &["0,1", "6"]),
    test_grouping_with_large_input => (r#"
        fun main() {
            val values = (1..12).toList()
            val grouped = values.groupBy { it % 4 }
            val totalFirst = grouped[0]!!.sum()
            val totalOther = grouped[1]!!.size + grouped[2]!!.size + grouped[3]!!.size
            println(totalFirst)
            println(totalOther)
        }
    "#, &["24", "9"]),
    test_grouping_by_remainder_and_join => (r#"
        fun main() {
            val grouped = listOf(5, 6, 7, 8, 9, 10).groupBy { it % 2 }
            val evenFirst = grouped[0]!![0]
            val oddLast = grouped[1]!![2]
            println(evenFirst + oddLast)
        }
    "#, &["16"]),
    test_grouping_zip_projection => (r#"
        fun main() {
            val grouped = listOf("aa", "bbb", "cccc", "d").groupBy { it.length }
            val keys = grouped.keys.sorted()
            val out = keys.joinToString(":")
            println(out)
        }
    "#, &["1:2:3:4"]),
    test_grouping_singleton_group => (r#"
        fun main() {
            val values = listOf("z")
            val grouped = values.groupBy { it }
            println(grouped["z"]!!.size)
            println(grouped.keys.size)
        }
    "#, &["1", "1"]),
    test_grouping_multiple_spaces => (r#"
        fun main() {
            val values = listOf("x y", "a b", "z z", "x y")
            val grouped = values.groupBy { it[0] }
            println(grouped['x']!!.size)
            println(grouped['a']!!.size)
        }
    "#, &["2", "1"]),
    test_grouping_after_mapping => (r#"
        fun main() {
            val words = listOf("one", "two", "three", "four")
            val grouped = words.map { it.uppercase() }.groupBy { it.length }
            println(grouped[3]!!.joinToString(","))
            println(grouped[4]!!.first())
        }
    "#, &["ONE,TWO", "FOUR"]) }
