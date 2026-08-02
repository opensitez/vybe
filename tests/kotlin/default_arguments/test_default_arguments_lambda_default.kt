// vybe-test: kotlin/default_arguments/test_default_arguments_lambda_default
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun apply(v: Int, op: (Int) -> Int = { it + 1 }): Int {
            return op(v)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((apply(4)).toString(), "5")
            __check((apply(4, { it * 2 })).toString(), "8")
        }
