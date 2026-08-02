// vybe-test: kotlin/infix/test_custom_infix_function
// origin: languages/kotlin/tests/kotlin/test_infix.rs

class Calculator(val base: Int) {
            infix fun plusValue(other: Int): Int {
                return base + other
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val calc = Calculator(10)
            val res = calc plusValue 5
            __check((res).toString(), "15")
        }
