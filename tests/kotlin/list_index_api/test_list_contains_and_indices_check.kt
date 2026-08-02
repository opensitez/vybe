// vybe-test: kotlin/list_index_api/test_list_contains_and_indices_check
// origin: languages/kotlin/tests/kotlin/test_list_index_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("x", "y", "z")
            __check((values.contains("y")).toString(), "true")
            __check((values.indices.first()).toString(), "0")
            __check((values.indices.last()).toString(), "2")
        }
