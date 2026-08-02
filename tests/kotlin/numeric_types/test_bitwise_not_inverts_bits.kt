// vybe-test: kotlin/numeric_types/test_bitwise_not_inverts_bits
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 0
            __check((value.inv()).toString(), "-1")
            __check(((-1).inv()).toString(), "0")
        }
