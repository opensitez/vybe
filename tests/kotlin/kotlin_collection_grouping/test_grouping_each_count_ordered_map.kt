// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_each_count_ordered_map
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("aa", "bb", "a", "cc", "d")
            val counts = values.groupingBy { it.length }.eachCount().toList().sortedBy { it.first }
            __check((counts.joinToString("|") { it.first.toString() + ":" + it.second }).toString(), "1:2|2:3")
        }
