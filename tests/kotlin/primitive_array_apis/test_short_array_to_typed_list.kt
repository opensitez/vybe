// vybe-test: kotlin/primitive_array_apis/test_short_array_to_typed_list
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = shortArrayOf(4, 5)
            __check((values.toTypedArray().joinToString(",")).toString(), "4,5")
            __check((values.size).toString(), "2")
        }
