// vybe-test: kotlin/nullability/test_nullable_with_elvis_and_binary_op
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a: Int? = null
            val b: Int? = 4
            val left = a ?: 0
            val right = b ?: 0
            __check((left + right).toString(), "4")
        }
