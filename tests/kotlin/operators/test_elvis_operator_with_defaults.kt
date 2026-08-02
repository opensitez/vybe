// vybe-test: kotlin/operators/test_elvis_operator_with_defaults
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val name: String? = null
            val provided: String? = "value"
            __check((name ?: "fallback").toString(), "fallback")
            __check((provided ?: "fallback").toString(), "value")
        }
