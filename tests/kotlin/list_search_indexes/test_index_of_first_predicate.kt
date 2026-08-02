// vybe-test: kotlin/list_search_indexes/test_index_of_first_predicate
// origin: languages/kotlin/tests/kotlin/test_list_search_indexes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(2, 4, 6, 7, 8)
            __check((values.indexOfFirst { it % 2 == 1 }).toString(), "3")
            __check((values.indexOfFirst { it > 5 }).toString(), "2")
        }
