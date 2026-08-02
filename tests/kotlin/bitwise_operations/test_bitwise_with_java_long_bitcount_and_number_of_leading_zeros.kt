// vybe-test: kotlin/bitwise_operations/test_bitwise_with_java_long_bitcount_and_number_of_leading_zeros
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val sample = 0b0001_0010
            __check((java.lang.Integer.bitCount(sample)).toString(), "2")
            __check((java.lang.Integer.numberOfLeadingZeros(sample)).toString(), "28")
            __check((java.lang.Integer.numberOfTrailingZeros(sample)).toString(), "1")
            __check((java.lang.Integer.numberOfTrailingZeros(0)).toString(), "32")
        }
