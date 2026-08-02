// vybe-test: kotlin/kotlin_operator_overflow/test_signed_shift_left_behavior
// origin: languages/kotlin/tests/kotlin/test_kotlin_operator_overflow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1 shl 30).toString(), "1073741824")
            __check((-1 shl 1).toString(), "-2")
        }
