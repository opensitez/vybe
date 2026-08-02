// vybe-test: kotlin/bitwise_operations/test_xor_with_positive_values
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((0b1111 xor 0b1010).toString(), "5")
            __check((0b1111 xor 0b1111).toString(), "0")
            __check((6 xor 3).toString(), "5")
            __check((0b1100 xor 0b1010).toString(), "6")
        }
