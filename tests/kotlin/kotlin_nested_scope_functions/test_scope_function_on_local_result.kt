// vybe-test: kotlin/kotlin_nested_scope_functions/test_scope_function_on_local_result
// origin: languages/kotlin/tests/kotlin/test_kotlin_nested_scope_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val n = with("ok") {
                length + 1
            }
            __check((n).toString(), "3")
        }
