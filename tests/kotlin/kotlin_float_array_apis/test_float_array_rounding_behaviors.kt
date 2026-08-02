// vybe-test: kotlin/kotlin_float_array_apis/test_float_array_rounding_behaviors
// origin: languages/kotlin/tests/kotlin/test_kotlin_float_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = floatArrayOf(1.2f, 2.7f)
            __check((a.sum()).toString(), "3.9")
        }
