// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_each_count_empty
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counts = emptyList<String>().groupingBy { it.length }.eachCount()
            __check((counts.isEmpty()).toString(), "true")
        }
