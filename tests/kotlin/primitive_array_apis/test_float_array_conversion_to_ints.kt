// vybe-test: kotlin/primitive_array_apis/test_float_array_conversion_to_ints
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = floatArrayOf(1.2f, 2.8f)
            val ints = values.map { it.toInt() }
            __check((ints.joinToString(",")).toString(), "1,2")
        }
