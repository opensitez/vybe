// vybe-test: kotlin/scope/test_scope_if_expression_has_its_own_binding_scope
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val label = "outer"
            val result = if (true) {
                val label = "then"
                label
            } else {
                val label = "else"
                label
            }

            __check((label).toString(), "outer")
            __check((result).toString(), "then")
        }
