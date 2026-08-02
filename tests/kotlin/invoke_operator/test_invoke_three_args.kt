// vybe-test: kotlin/invoke_operator/test_invoke_three_args
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Sum3 {
            operator fun invoke(a: Int, b: Int, c: Int): Int = a + b + c
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Sum3()(1, 2, 3)).toString(), "6")
        }
