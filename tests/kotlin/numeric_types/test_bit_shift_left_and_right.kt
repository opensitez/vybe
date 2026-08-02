// vybe-test: kotlin/numeric_types/test_bit_shift_left_and_right
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 1
            __check((value shl 4).toString(), "16")
            __check((16 shr 2).toString(), "4")
            __check((-16 shr 2).toString(), "-4")
        }
