// vybe-test: kotlin/bitwise_operations/test_shift_right_arithmetic_negative
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((-8 shr 1).toString(), "-4")
            __check((-8 shr 2).toString(), "-2")
            __check((-1 shr 3).toString(), "-1")
            __check((-15 shr 1).toString(), "-8")
        }
