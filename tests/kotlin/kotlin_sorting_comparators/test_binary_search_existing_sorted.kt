// vybe-test: kotlin/kotlin_sorting_comparators/test_binary_search_existing_sorted
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 3, 5, 7)
            __check((values.binarySearch(5)).toString(), "2")
            __check((values.binarySearch(6)).toString(), "-4")
        }
