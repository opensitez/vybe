// vybe-test: kotlin/bitwise_operations/test_and_with_positive_values
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((0b1111 and 0b1010).toString(), "10")
            __check((0b0101 and 0b1010).toString(), "0")
            __check((12 and 10).toString(), "8")
            __check((4 and 3).toString(), "0")
        }
