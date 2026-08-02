// vybe-test: kotlin/operators/test_custom_unary_operators
// origin: languages/kotlin/tests/kotlin/test_operators.rs

class Flag(val value: Int) {
            operator fun unaryMinus(): Flag = Flag(-value)
            operator fun unaryPlus(): Flag = Flag(+value)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Flag(8)
            __check(((-value).value).toString(), "-8")
            __check(((+value).value).toString(), "8")
        }
