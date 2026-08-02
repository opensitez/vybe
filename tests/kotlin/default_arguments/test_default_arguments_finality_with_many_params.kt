// vybe-test: kotlin/default_arguments/test_default_arguments_finality_with_many_params
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun eval(a: Int, b: Int = 1, c: Int = 2, d: Int = 3): Int {
            return a + b + c + d
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((eval(1)).toString(), "7")
            __check((eval(1, d = 10)).toString(), "16")
            __check((eval(1, 2, 3, 4)).toString(), "10")
        }
