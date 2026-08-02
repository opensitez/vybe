// vybe-test: kotlin/kotlin_nested_scope_functions/test_nested_scope_with_if_expression
// origin: languages/kotlin/tests/kotlin/test_kotlin_nested_scope_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = run {
                if (2 > 1) {
                    "yes"
                } else {
                    "no"
                }
            }
            __check((out).toString(), "yes")
        }
