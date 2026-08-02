// vybe-test: kotlin/lambdas/test_no_arg_lambda
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val sayHi = { "Hi" }
            __check((sayHi()).toString(), "Hi")
        }
