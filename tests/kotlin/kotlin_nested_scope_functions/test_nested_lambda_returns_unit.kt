// vybe-test: kotlin/kotlin_nested_scope_functions/test_nested_lambda_returns_unit
// origin: languages/kotlin/tests/kotlin/test_kotlin_nested_scope_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val printer = { text: String ->
                __check((text).toString(), "ok")
                Unit
            }
            printer("ok")
            __check(("done").toString(), "done")
        }
