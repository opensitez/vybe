// vybe-test: kotlin/kotlin_big_numbers/test_big_integer_division_and_remainder
// origin: languages/kotlin/tests/kotlin/test_kotlin_big_numbers.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = java.math.BigInteger("100")
            val b = java.math.BigInteger("9")
            __check((a.divide(b).toString()).toString(), "11")
            __check((a.remainder(b).toString()).toString(), "1")
            val q = a.divideAndRemainder(b)
            __check((q[0].toString()).toString(), "11")
            __check((q[1].toString()).toString(), "1")
        }
