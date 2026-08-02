// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_each_count
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("aa", "bb", "a", "bb", "a", "a")
            val counts = values.groupingBy { it.length }.eachCount()
            __check((counts[1]).toString(), "1")
            __check((counts[2]).toString(), "5")
        }
