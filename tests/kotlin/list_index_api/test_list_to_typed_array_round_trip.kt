// vybe-test: kotlin/list_index_api/test_list_to_typed_array_round_trip
// origin: languages/kotlin/tests/kotlin/test_list_index_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(4, 5, 6)
            val arr = values.toTypedArray()
            val back = arr.toList()
            __check((back.size).toString(), "3")
            __check((back.joinToString(",")).toString(), "4,5,6")
            __check(((back === values).toString()).toString(), "false")
        }
