// vybe-test: kotlin/infix/test_infix_right_operand_is_full_expression
// origin: languages/kotlin/tests/kotlin/test_infix.rs

class Score(val value: Int) {
            infix fun plus(other: Int): Int = value + other
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val calc = Score(2)
            val total = (calc plus 3) * 4
            __check((total).toString(), "20")
        }
