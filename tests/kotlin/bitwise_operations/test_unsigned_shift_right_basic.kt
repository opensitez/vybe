// vybe-test: kotlin/bitwise_operations/test_unsigned_shift_right_basic
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((-1 ushr 1).toString(), "2147483647")
            __check((-16 ushr 2).toString(), "1073741820")
            __check((16 ushr 1).toString(), "8")
            __check((1 ushr 1).toString(), "0")
        }
