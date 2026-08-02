// vybe-test: kotlin/scoping_functions/test_let_shadowed_name_does_not_escape_scope
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "outer"
            val projected = value.let { value ->
                value.uppercase()
            }
            __check((projected).toString(), "OUTER")
            __check((value).toString(), "outer")
        }
