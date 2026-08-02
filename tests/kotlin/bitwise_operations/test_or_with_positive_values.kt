// vybe-test: kotlin/bitwise_operations/test_or_with_positive_values
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((0b1111 or 0b1010).toString(), "15")
            __check((0b0001 or 0b0010).toString(), "3")
            __check((4 or 3).toString(), "7")
            __check((8 or 1).toString(), "9")
        }
