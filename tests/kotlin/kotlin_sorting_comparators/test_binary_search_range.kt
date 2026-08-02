// vybe-test: kotlin/kotlin_sorting_comparators/test_binary_search_range
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 3, 5, 7, 9)
            __check((values.binarySearch(5, 0, 2)).toString(), "-3")
            __check((values.binarySearch(3, 0, 2)).toString(), "1")
        }
