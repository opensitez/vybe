// vybe-test: kotlin/primitive_array_apis/test_ullong_array_to_string
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = ulongArrayOf(1UL, 2UL)
            __check((values.joinToString(",")).toString(), "1,2")
            __check((values[1]).toString(), "2")
        }
