// vybe-test: kotlin/kotlin_nested_scope_functions/test_nested_lambda_captures_outer_variable
// origin: languages/kotlin/tests/kotlin/test_kotlin_nested_scope_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var total = 1
            val inc = {
                val local = 2
                total += local
            }
            inc()
            __check((total).toString(), "3")
        }
