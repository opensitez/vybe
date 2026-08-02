// vybe-test: kotlin/invoke_operator/test_invoke_array_argument
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Adder {
            operator fun invoke(values: IntArray): Int = values.sum()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Adder()(intArrayOf(3, 4, 5))).toString(), "12")
        }
