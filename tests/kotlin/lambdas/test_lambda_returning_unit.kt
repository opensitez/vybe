// vybe-test: kotlin/lambdas/test_lambda_returning_unit
// origin: languages/kotlin/tests/kotlin/test_lambdas.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val action: (String) -> Unit = { s -> __check((s).toString(), "go") }
action("go") }
