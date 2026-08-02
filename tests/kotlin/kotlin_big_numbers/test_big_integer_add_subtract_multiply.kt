// vybe-test: kotlin/kotlin_big_numbers/test_big_integer_add_subtract_multiply
// origin: languages/kotlin/tests/kotlin/test_kotlin_big_numbers.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = java.math.BigInteger("12345678901234567890")
            val b = java.math.BigInteger("987654321")
            __check((a.add(b).toString()).toString(), "12345679888888888891")
            __check((a.subtract(b).toString()).toString(), "12345677913580246769")
            __check((a.multiply(java.math.BigInteger("2")).toString()).toString(), "24691357802469135680")
        }
