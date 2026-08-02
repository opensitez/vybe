// vybe-test: kotlin/primitive_array_apis/test_byte_array_sum_and_average
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = byteArrayOf(1, 2, 3)
            val sum = values.sum()
            val avg = values.average()
            __check((sum).toString(), "6")
            __check((avg).toString(), "2.0")
        }
