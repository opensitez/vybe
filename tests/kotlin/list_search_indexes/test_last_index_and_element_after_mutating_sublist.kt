// vybe-test: kotlin/list_search_indexes/test_last_index_and_element_after_mutating_sublist
// origin: languages/kotlin/tests/kotlin/test_list_search_indexes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 2, 3, 4)
            val window = values.subList(1, 3)
            __check((window.lastIndex).toString(), "1")
            window.clear()
            __check((values.joinToString(",")).toString(), "1,4")
            __check((values.lastIndex).toString(), "1")
        }
