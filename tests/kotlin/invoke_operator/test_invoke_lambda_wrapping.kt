// vybe-test: kotlin/invoke_operator/test_invoke_lambda_wrapping
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Wrapper {
            operator fun invoke(v: (Int) -> Int): Int = v(6)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val w = Wrapper()
            __check((w { it * 2 }).toString(), "12")
        }
