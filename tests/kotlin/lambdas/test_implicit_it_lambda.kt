// vybe-test: kotlin/lambdas/test_implicit_it_lambda
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val doubleIt = { it * 2 }
            __check((doubleIt(21)).toString(), "42")
        }
