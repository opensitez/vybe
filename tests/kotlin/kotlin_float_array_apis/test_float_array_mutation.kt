// vybe-test: kotlin/kotlin_float_array_apis/test_float_array_mutation
// origin: languages/kotlin/tests/kotlin/test_kotlin_float_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = FloatArray(2)
            a[0] = 3.5f
            a[1] = a[0] * 2
            __check((a[1]).toString(), "7.0")
        }
