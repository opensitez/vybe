// vybe-test: kotlin/scope/test_outer_scope_after_lambda_call
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "outer"
            val action = {
                val value = "inner"
                __check((value).toString(), "inner")
            }
            action()
            __check((value).toString(), "outer")
        }
