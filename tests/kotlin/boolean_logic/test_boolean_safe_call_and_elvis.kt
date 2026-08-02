// vybe-test: kotlin/boolean_logic/test_boolean_safe_call_and_elvis
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Boolean? = null
            __check((value?.toString() ?: "missing").toString(), "missing")
            val another: Boolean? = false
            __check((another?.toString() ?: "missing").toString(), "false")
        }
