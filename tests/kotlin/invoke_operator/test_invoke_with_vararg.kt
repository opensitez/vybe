// vybe-test: kotlin/invoke_operator/test_invoke_with_vararg
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Summer {
            operator fun invoke(vararg values: Int): Int = values.sum()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Summer()(1, 2, 3)).toString(), "6")
        }
