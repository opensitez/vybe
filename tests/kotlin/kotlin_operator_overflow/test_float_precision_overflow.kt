// vybe-test: kotlin/kotlin_operator_overflow/test_float_precision_overflow
// origin: languages/kotlin/tests/kotlin/test_kotlin_operator_overflow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = Float.MAX_VALUE
            __check(((x * 2).isInfinite()).toString(), "true")
        }
