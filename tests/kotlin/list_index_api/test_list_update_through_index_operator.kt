// vybe-test: kotlin/list_index_api/test_list_update_through_index_operator
// origin: languages/kotlin/tests/kotlin/test_list_index_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 2, 3)
            values[0] = 7
            values[2] = 9
            __check((values.joinToString(",")).toString(), "7,2,9")
        }
