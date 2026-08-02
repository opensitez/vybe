// vybe-test: kotlin/kotlin_number_conversion_apis/test_nan_and_infinity_parse
// origin: languages/kotlin/tests/kotlin/test_kotlin_number_conversion_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nan = 0.0 / 0.0
            val inf = 1.0 / 0.0
            val ninf = -1.0 / 0.0
            __check((nan.isNaN()).toString(), "true")
            __check((inf.isInfinite()).toString(), "true")
            __check((ninf.isInfinite()).toString(), "true")
            __check(((inf + inf).isInfinite()).toString(), "true")
        }
