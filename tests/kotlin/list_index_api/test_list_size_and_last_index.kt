// vybe-test: kotlin/list_index_api/test_list_size_and_last_index
// origin: languages/kotlin/tests/kotlin/test_list_index_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(10, 20, 30, 40)
            __check((values.size).toString(), "4")
            __check((values.lastIndex).toString(), "3")
        }
