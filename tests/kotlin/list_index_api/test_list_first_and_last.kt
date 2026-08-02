// vybe-test: kotlin/list_index_api/test_list_first_and_last
// origin: languages/kotlin/tests/kotlin/test_list_index_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(3, 5, 8)
            __check((values.first()).toString(), "3")
            __check((values.last()).toString(), "8")
            __check((values.firstOrNull()).toString(), "3")
            __check((values.lastOrNull()).toString(), "8")
        }
