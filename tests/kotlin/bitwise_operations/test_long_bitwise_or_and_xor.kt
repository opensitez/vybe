// vybe-test: kotlin/bitwise_operations/test_long_bitwise_or_and_xor
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Long = 0b1010
            val next: Long = 0b1100
            __check((value or next).toString(), "14")
            __check((value and next).toString(), "8")
            __check((value xor next).toString(), "6")
        }
