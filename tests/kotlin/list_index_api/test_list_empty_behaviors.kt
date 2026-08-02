// vybe-test: kotlin/list_index_api/test_list_empty_behaviors
// origin: languages/kotlin/tests/kotlin/test_list_index_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = emptyList<Int>()
            __check((values.isEmpty()).toString(), "true")
            __check((values.firstOrNull() ?: "none").toString(), "none")
            __check((values.lastOrNull() ?: "none").toString(), "none")
            __check((values.elementAtOrElse(0) { it + 4 }).toString(), "4")
            __check((values.elementAtOrNull(0) ?: "none").toString(), "none")
        }
