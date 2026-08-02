// vybe-test: kotlin/numeric_types/test_bitwise_or_and_xor
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 0b1110
            val b = 0b0110
            __check((a or b).toString(), "14")
            __check((a xor b).toString(), "8")
        }
