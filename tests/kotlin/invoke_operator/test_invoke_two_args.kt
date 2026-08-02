// vybe-test: kotlin/invoke_operator/test_invoke_two_args
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class PairAdder {
            operator fun invoke(a: Int, b: Int): Int = a + b
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((PairAdder()(2, 5)).toString(), "7")
        }
