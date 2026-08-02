// vybe-test: kotlin/primitive_array_apis/test_int_array_copy_of_range
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val src = intArrayOf(1, 2, 3, 4)
            val dst = src.copyOfRange(1, 3)
            __check((dst.joinToString(",")).toString(), "2,3")
        }
