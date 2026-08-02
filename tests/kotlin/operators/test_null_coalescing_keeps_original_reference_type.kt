// vybe-test: kotlin/operators/test_null_coalescing_keeps_original_reference_type
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source: Any? = "value"
            val text: String = source as? String ?: "fallback"
            __check((text).toString(), "value")
            val raw: Any? = null
            val again: String = raw as? String ?: "fallback"
            __check((again).toString(), "fallback")
        }
