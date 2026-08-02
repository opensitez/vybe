// vybe-test: kotlin/boolean_logic/test_boolean_let_and_scope_functions
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Boolean? = true
            val transformed = value?.let { it && true } ?: false
            __check((transformed).toString(), "true")
            val none: Boolean? = null
            __check((none?.let { it } ?: false).toString(), "false")
        }
