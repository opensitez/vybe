// vybe-test: kotlin/bitwise_operations/test_shift_count_wraps_for_int
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1 shl 31).toString(), "-2147483648")
            __check((1 shl 32).toString(), "1")
            __check((1 shl 40).toString(), "256")
            __check((1 shl -1).toString(), "-2147483648")
        }
