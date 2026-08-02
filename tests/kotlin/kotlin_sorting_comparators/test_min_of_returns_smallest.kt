// vybe-test: kotlin/kotlin_sorting_comparators/test_min_of_returns_smallest
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(7, 3, 9, 1)
            __check((values.minOrNull()).toString(), "1")
            __check((values.maxOrNull()).toString(), "9")
        }
