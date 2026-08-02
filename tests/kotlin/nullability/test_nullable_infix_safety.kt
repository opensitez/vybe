// vybe-test: kotlin/nullability/test_nullable_infix_safety
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left: Int? = 3
            val right: Int? = null
            val leftValue = left ?: 0
            val rightValue = right ?: 10
            __check((leftValue + rightValue).toString(), "13")
        }
