// vybe-test: kotlin/list_search_indexes/test_last_index_for_duplicates_and_absent_value
// origin: languages/kotlin/tests/kotlin/test_list_search_indexes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 4, 4, 4, 9)
            __check((values.lastIndexOf(4)).toString(), "3")
            __check((values.lastIndexOf(2)).toString(), "-1")
        }
