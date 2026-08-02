// vybe-test: kotlin/bitwise_operations/test_bitwise_precedence_with_arithmetic
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1 or 2 + 4 and 8).toString(), "9")
            __check(((1 or 2) + (4 and 8)).toString(), "3")
            __check((2 shl 3 + 1).toString(), "16")
            __check((2 shl (3 + 1)).toString(), "32")
        }
