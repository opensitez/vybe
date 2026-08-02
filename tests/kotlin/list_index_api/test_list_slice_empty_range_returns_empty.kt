// vybe-test: kotlin/list_index_api/test_list_slice_empty_range_returns_empty
// origin: languages/kotlin/tests/kotlin/test_list_index_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 3, 4)
            val part = values.slice(2 until 2)
            __check((part.isEmpty()).toString(), "true")
            __check((part.size).toString(), "0")
        }
