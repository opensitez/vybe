// vybe-test: kotlin/kotlin_operator_overflow/test_bitwise_and_masks
// origin: languages/kotlin/tests/kotlin/test_kotlin_operator_overflow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((0b1010 and 0b0111).toString(), "2")
            __check((0b1010 or 0b0101).toString(), "15")
            __check((0b1010 xor 0b1111).toString(), "5")
        }
