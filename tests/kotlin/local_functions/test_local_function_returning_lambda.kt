// vybe-test: kotlin/local_functions/test_local_function_returning_lambda
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun make(prefix: String): (Int) -> String {
                return { value -> "$prefix$value" }
            }
            val f = make("x")
            __check((f(9)).toString(), "x9")
        }
