// vybe-test: kotlin/invoke_operator/test_invoke_one_arg
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Math {
            operator fun invoke(v: Int): Int = v + 1
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Math()(3)).toString(), "4")
        }
