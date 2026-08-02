// vybe-test: kotlin/bitwise_operations/test_bitwise_precedes_comparison_logic
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val raw = 0b1010
            __check(((raw and 2) == 2).toString(), "true")
            __check(((raw and 1) == 1).toString(), "false")
            __check(((raw or 1) > raw).toString(), "true")
            __check(((raw xor 0) == raw).toString(), "true")
        }
