// vybe-test: kotlin/kotlin_sorting_comparators/test_sorted_numbers_default_order
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(4, 1, 3, 2)
            __check((values.sorted().joinToString(",")).toString(), "1,2,3,4")
        }
