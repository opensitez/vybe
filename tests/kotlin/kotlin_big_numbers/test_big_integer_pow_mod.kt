// vybe-test: kotlin/kotlin_big_numbers/test_big_integer_pow_mod
// origin: languages/kotlin/tests/kotlin/test_kotlin_big_numbers.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = java.math.BigInteger("2")
            val p = x.pow(10)
            __check((p.toString()).toString(), "1024")
            val m = p.mod(java.math.BigInteger("1000"))
            __check((m.toString()).toString(), "24")
            __check((x.modPow(java.math.BigInteger("5"), java.math.BigInteger("13")).toString()).toString(), "6")
        }
