// vybe-test: kotlin/operators/test_nested_elvis_chain_and_rhs_evaluation_boundary
// origin: languages/kotlin/tests/kotlin/test_operators.rs

var fallbackCalls = 0

        fun fallback(value: String?): String {
            fallbackCalls += 1
            return value ?: "default"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first: String? = null
            val second: String? = null
            val third: String? = "value"
            val present: String? = "keep"
            __check((first ?: second ?: fallback(third)).toString(), "value")
            __check((fallbackCalls).toString(), "1")
            fallbackCalls = 0
            __check((present ?: fallback(present)).toString(), "keep")
            __check((fallbackCalls).toString(), "0")
        }
