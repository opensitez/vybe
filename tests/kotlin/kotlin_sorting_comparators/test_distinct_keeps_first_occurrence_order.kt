// vybe-test: kotlin/kotlin_sorting_comparators/test_distinct_keeps_first_occurrence_order
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 2, 3, 1, 4, 3)
            __check((values.distinct().joinToString(",")).toString(), "1,2,3,4")
        }
