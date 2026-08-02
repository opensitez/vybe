// vybe-test: kotlin/kotlin_nested_scope_functions/test_local_function_uses_outer_arguments
// origin: languages/kotlin/tests/kotlin/test_kotlin_nested_scope_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun format(v: Int): String {
                return "v=${'$'}v"
            }
            __check((format(7)).toString(), "v=7")
        }
