// vybe-test: kotlin/kotlin_sorting_comparators/test_sorted_descending_numbers
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 4, 2, 3)
            __check((values.sortedDescending().joinToString(",")).toString(), "4,3,2,1")
        }
