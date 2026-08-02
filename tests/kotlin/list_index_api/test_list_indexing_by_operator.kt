// vybe-test: kotlin/list_index_api/test_list_indexing_by_operator
// origin: languages/kotlin/tests/kotlin/test_list_index_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf("a", "b", "c")
            __check((values[0]).toString(), "a")
            __check((values[1]).toString(), "b")
            __check((values[2]).toString(), "c")
        }
