// vybe-test: kotlin/numeric_types/test_bitwise_and_masking
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 0b1111 and 0b1010
            val mask = 0b0101 and value
            __check((value).toString(), "10")
            __check((mask).toString(), "0")
        }
