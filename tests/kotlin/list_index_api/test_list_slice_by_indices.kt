// vybe-test: kotlin/list_index_api/test_list_slice_by_indices
// origin: languages/kotlin/tests/kotlin/test_list_index_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(7, 8, 9, 10, 11)
            val part = values.slice(1..3)
            __check((part.joinToString(",")).toString(), "8,9,10")
            __check((part.size).toString(), "3")
        }
