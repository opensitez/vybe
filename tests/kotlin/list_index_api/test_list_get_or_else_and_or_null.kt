// vybe-test: kotlin/list_index_api/test_list_get_or_else_and_or_null
// origin: languages/kotlin/tests/kotlin/test_list_index_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(4, 5, 6)
            __check((values.getOrElse(1) { 0 }).toString(), "5")
            __check((values.getOrElse(4) { 99 }).toString(), "99")
            __check((values.getOrNull(0) ?: -1).toString(), "4")
            __check((values.getOrNull(4) ?: -1).toString(), "-1")
        }
