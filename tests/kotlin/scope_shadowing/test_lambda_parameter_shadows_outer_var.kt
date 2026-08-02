// vybe-test: kotlin/scope_shadowing/test_lambda_parameter_shadows_outer_var
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "outer"
            val transform = { value: Int -> "lambda:$value" }
            __check((transform(7)).toString(), "lambda:7")
            __check((value).toString(), "outer")
        }
