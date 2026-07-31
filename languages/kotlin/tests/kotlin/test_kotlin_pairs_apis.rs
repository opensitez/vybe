kotlin_run_cases! {
    test_pair_infix_constructor => (r#"
        fun main() {
            val value = "left" to "right"
            println(value.first)
            println(value.second)
        }
    "#, &["left", "right"]),
    test_pair_explicit_constructor => (r#"
        fun main() {
            val value = Pair(1, 2)
            println(value.first)
            println(value.second)
        }
    "#, &["1", "2"]),
    test_pair_destructure => (r#"
        fun main() {
            val (first, second) = "a" to "b"
            println(first)
            println(second)
        }
    "#, &["a", "b"]),
    test_pair_component_functions => (r#"
        fun main() {
            val value = 7 to 9
            println(value.component1())
            println(value.component2())
        }
    "#, &["7", "9"]),
    test_pair_equality_same_values => (r#"
        fun main() {
            val a = 4 to 5
            val b = 4 to 5
            println(a == b)
            println(a != b)
        }
    "#, &["true", "false"]),
    test_pair_inequality => (r#"
        fun main() {
            val a = 1 to 2
            val b = 2 to 1
            println(a == b)
        }
    "#, &["false"]),
    test_pair_to_string => (r#"
        fun main() {
            val p = "x" to 1
            println(p.toString())
        }
    "#, &["(x, 1)"]),
    test_pair_in_map_entries => (r#"
        fun main() {
            val values = mapOf(1 to "one", 2 to "two")
            var out = ""
            for (entry in values) {
                out += entry.key.toString() + entry.value
            }
            println(out)
        }
    "#, &["1one2two"]),
    test_pair_list_destructure_loop => (r#"
        fun main() {
            val pairs = listOf("a" to 1, "b" to 2)
            var out = ""
            for ((k, v) in pairs) {
                out += k + v.toString()
            }
            println(out)
        }
    "#, &["a1b2"]),
    test_pair_map_from_list => (r#"
        fun main() {
            val source = listOf("a" to 1, "b" to 2)
            val map = source.toMap()
            println(map["a"])
            println(map["b"])
        }
    "#, &["1", "2"]),
    test_pair_duplicate_key_keeps_last => (r#"
        fun main() {
            val source = listOf("x" to 1, "x" to 9)
            val map = source.toMap()
            println(map["x"])
            println(map.size)
        }
    "#, &["9", "1"]),
    test_pair_zip_basic => (r#"
        fun main() {
            val zipped = listOf(1, 2, 3).zip(listOf("a", "b", "c"))
            println(zipped[0].first)
            println(zipped[1].second)
            println(zipped.joinToString("|") { it.toString() })
        }
    "#, &["1", "b", "(1, a)|(2, b)|(3, c)"]),
    test_pair_zip_with_transform => (r#"
        fun main() {
            val merged = listOf(1, 2, 3).zip(listOf("a", "b", "c")) { i, s -> s + i }
            println(merged.joinToString(","))
        }
    "#, &["a1,b2,c3"]),
    test_pair_unzip_string => (r#"
        fun main() {
            val values = listOf("x" to 1, "y" to 2, "z" to 3)
            val (letters, numbers) = values.unzip()
            println(letters.joinToString(""))
            println(numbers.joinToString(","))
        }
    "#, &["xyz", "1,2,3"]),
    test_pair_unzip_empty => (r#"
        fun main() {
            val (a, b) = emptyList<Pair<Int, Int>>().unzip()
            println(a.isEmpty())
            println(b.isEmpty())
        }
    "#, &["true", "true"]),
    test_triple_accessors => (r#"
        fun main() {
            val t = Triple("alpha", 11, true)
            println(t.first)
            println(t.second)
            println(t.third)
        }
    "#, &["alpha", "11", "true"]),
    test_triple_destructure => (r#"
        fun main() {
            val (a, b, c) = Triple(2, 4, 6)
            println(a)
            println(b)
            println(c)
        }
    "#, &["2", "4", "6"]),
    test_triple_equal_same_values => (r#"
        fun main() {
            val t1 = Triple(1, 2, 3)
            val t2 = Triple(1, 2, 3)
            println(t1 == t2)
            println(t1.hashCode() == t2.hashCode())
        }
    "#, &["true", "true"]),
    test_triple_not_equal => (r#"
        fun main() {
            val t1 = Triple(1, 2, 3)
            val t2 = Triple(3, 2, 1)
            println(t1 == t2)
        }
    "#, &["false"]),
    test_triple_string_form => (r#"
        fun main() {
            val t = Triple("k", "t", "v")
            println(t.toString())
        }
    "#, &["(k, t, v)"]),
    test_triple_in_map => (r#"
        fun main() {
            val values = mapOf("row" to Triple(1, 2, 3), "other" to Triple(4, 5, 6))
            println(values["row"]?.first)
            println(values["other"]?.third)
        }
    "#, &["1", "6"]),
    test_pair_and_triple_nested => (r#"
        fun main() {
            val nested = Pair("outer", Triple(1, "mid", 3))
            val (left, right) = nested
            println(left)
            println(right.second)
        }
    "#, &["outer", "mid"]),
    test_pair_with_char_digit => (r#"
        fun main() {
            val value = 'A' to 5
            println(value.first.code)
            println(value.second)
            println(value.toString())
        }
    "#, &["65", "5", "(A, 5)"]),
    test_pair_negative_numbers => (r#"
        fun main() {
            val values = listOf(-1 to -2, 3 to 4)
            val total = values.map { it.first + it.second }.joinToString(",")
            println(total)
            println(values[0].first < values[1].first)
        }
    "#, &["-3,7", "true"]),
    test_pair_with_boolean => (r#"
        fun main() {
            val values = listOf(true to "on", false to "off")
            var out = ""
            for ((state, name) in values) {
                out += if (state) name.uppercase() else name
            }
            println(out)
        }
    "#, &["ONoff"]),
    test_pair_of_empty_string => (r#"
        fun main() {
            val p = "" to 0
            println(p.first.isEmpty())
            println(p.second)
        }
    "#, &["true", "0"]),
    test_pair_sort_by_first => (r#"
        fun main() {
            val values = listOf(5 to "five", 2 to "two", 4 to "four")
            val sorted = values.sortedBy { it.first }
            println(sorted.joinToString("|") { it.first.toString() })
        }
    "#, &["2|4|5"]),
    test_pair_sort_by_second => (r#"
        fun main() {
            val values = listOf("cat" to 3, "bee" to 1, "ant" to 2)
            val sorted = values.sortedBy { it.second }
            println(sorted.joinToString("|") { it.first.toString() })
        }
    "#, &["bee|ant|cat"]),
    test_pair_sum_projection => (r#"
        fun main() {
            val values = listOf(1 to 10, 2 to 20, 3 to 30)
            val total = values.map { it.first + it.second }.sum()
            println(total)
        }
    "#, &["66"]),
    test_pair_filter_by_second => (r#"
        fun main() {
            val values = listOf(1 to 10, 2 to 3, 3 to 8)
            val filtered = values.filter { it.second > 4 }
            println(filtered.joinToString(",") { it.first.toString() })
        }
    "#, &["1,3"]),
    test_pair_any_even_first => (r#"
        fun main() {
            val values = listOf(1 to 0, 2 to 9, 3 to 8)
            println(values.any { it.first % 2 == 0 })
            println(values.none { it.second < 0 })
        }
    "#, &["true", "true"]),
}
