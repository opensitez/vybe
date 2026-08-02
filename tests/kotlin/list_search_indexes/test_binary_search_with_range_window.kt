// vybe-test: kotlin/list_search_indexes/test_binary_search_with_range_window
// origin: languages/kotlin/tests/kotlin/test_list_search_indexes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 3, 5, 7, 9)
            __check((values.binarySearch(5, 1, 4)).toString(), "2")
            __check((values.binarySearch(5, 0, 2)).toString(), "-2")
        }
