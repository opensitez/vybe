// vybe-test: kotlin/primitive_array_apis/test_int_array_index_assignment
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = IntArray(3)
            values[1] = 9
            __check((values[0]).toString(), "0")
            __check((values[1]).toString(), "9")
            __check((values[2]).toString(), "0")
        }
