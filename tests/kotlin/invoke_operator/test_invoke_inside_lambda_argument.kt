// vybe-test: kotlin/invoke_operator/test_invoke_inside_lambda_argument
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Twice {
            operator fun invoke(v: Int): Int = v * 2
        }
        fun run(v: Int, fn: (Int) -> Int): Int = fn(v)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((run(5, Twice())).toString(), "10")
        }
