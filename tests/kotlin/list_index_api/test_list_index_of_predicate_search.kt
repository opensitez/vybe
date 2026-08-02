// vybe-test: kotlin/list_index_api/test_list_index_of_predicate_search
// origin: languages/kotlin/tests/kotlin/test_list_index_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 3, 2, 4)
            __check((values.indexOf(2)).toString(), "1")
            __check((values.lastIndexOf(2)).toString(), "3")
            __check((values.indexOfFirst { it > 2 }).toString(), "2")
            __check((values.indexOfLast { it % 2 == 0 }).toString(), "3")
        }
