// vybe-test: kotlin/kotlin_return_expressions/test_try_as_expression
// origin: languages/kotlin/tests/kotlin/test_kotlin_return_expressions.rs

fun safe(v: Int): Int = try {
            val inv = 10 / v
            inv
        } catch (e: ArithmeticException) {
            0
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((safe(2)).toString(), "5")
            __check((safe(0)).toString(), "0")
        }
