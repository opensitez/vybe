// vybe-test: kotlin/boolean_logic/test_boolean_nullable_coercion
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Boolean? = null
            val fallback = value ?: false
            __check((fallback).toString(), "false")
            val value2: Boolean? = true
            __check((value2 ?: false).toString(), "true")
        }
