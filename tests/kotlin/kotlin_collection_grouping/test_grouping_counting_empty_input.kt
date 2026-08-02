// vybe-test: kotlin/kotlin_collection_grouping/test_grouping_counting_empty_input
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_grouping.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val grouped = emptyList<Int>().groupBy { it }
            __check((grouped.isEmpty()).toString(), "true")
        }
