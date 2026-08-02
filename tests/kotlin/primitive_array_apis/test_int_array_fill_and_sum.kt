// vybe-test: kotlin/primitive_array_apis/test_int_array_fill_and_sum
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = IntArray(3) { it + 1 }
            __check((values.joinToString(",")).toString(), "1,2,3")
            __check((values.sum()).toString(), "6")
        }
