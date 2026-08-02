// vybe-test: kotlin/bitwise_operations/test_shift_right_arithmetic_positive
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((16 shr 1).toString(), "8")
            __check((16 shr 4).toString(), "1")
            __check((16 shr 5).toString(), "0")
            __check((3 shr 1).toString(), "1")
        }
