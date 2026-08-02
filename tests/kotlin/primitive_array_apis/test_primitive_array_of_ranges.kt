// vybe-test: kotlin/primitive_array_apis/test_primitive_array_of_ranges
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = IntArray(5) { it * it }
            __check((values.first()).toString(), "0")
            __check((values.last()).toString(), "16")
            __check((values[2]).toString(), "4")
        }
