// vybe-test: kotlin/list_search_indexes/test_binary_search_on_sorted_int_list
// origin: languages/kotlin/tests/kotlin/test_list_search_indexes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 3, 5, 7, 9)
            __check((values.binarySearch(5)).toString(), "2")
            __check((values.binarySearch(6)).toString(), "-3")
        }
