// vybe-test: kotlin/bitwise_operations/test_shift_left_basic
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1 shl 0).toString(), "1")
            __check((1 shl 1).toString(), "2")
            __check((1 shl 4).toString(), "16")
            __check((3 shl 2).toString(), "12")
        }
