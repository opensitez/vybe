// vybe-test: kotlin/primitive_array_apis/test_array_slice_copy
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = intArrayOf(10, 20, 30, 40, 50)
            __check((values.sliceArray(1..3).joinToString(",")).toString(), "20,30,40")
            __check((values.slice(1..3).joinToString(",")).toString(), "20,30,40")
        }
