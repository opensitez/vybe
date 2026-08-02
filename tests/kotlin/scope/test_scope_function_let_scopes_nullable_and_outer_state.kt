// vybe-test: kotlin/scope/test_scope_function_let_scopes_nullable_and_outer_state
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val prefix = "x"
            val result = prefix.let {
                val suffix = it.toUpperCase()
                suffix + "!"
            }
            __check((result).toString(), "X!")
            __check((prefix).toString(), "x")
        }
