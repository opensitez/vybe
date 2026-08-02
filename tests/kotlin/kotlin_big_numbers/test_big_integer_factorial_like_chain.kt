// vybe-test: kotlin/kotlin_big_numbers/test_big_integer_factorial_like_chain
// origin: languages/kotlin/tests/kotlin/test_kotlin_big_numbers.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val one = java.math.BigInteger.ONE
            val two = java.math.BigInteger("2")
            val three = java.math.BigInteger("3")
            val product = one.multiply(two).multiply(three)
            __check((product.toString()).toString(), "6")
            __check((product == java.math.BigInteger("6")).toString(), "true")
        }
