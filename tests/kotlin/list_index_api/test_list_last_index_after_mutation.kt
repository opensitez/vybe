// vybe-test: kotlin/list_index_api/test_list_last_index_after_mutation
// origin: languages/kotlin/tests/kotlin/test_list_index_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 2)
            values.add(3)
            __check((values.lastIndex).toString(), "2")
            values.removeAt(2)
            __check((values.lastIndex).toString(), "1")
        }
