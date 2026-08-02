// vybe-test: kotlin/scope/test_lambda_parameter_shadowing_outer_name
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var value = 7
            val toString = { value: Int -> value + 1 }
            __check((value).toString(), "7")
            __check((toString(3)).toString(), "4")
            __check((value).toString(), "7")
        }
