// vybe-test: kotlin/evaluation_order/test_elvis_uses_rhs_only_when_lhs_null
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val lhs: String? = null
            var order = ""
            val rhs = run { order += "rhs"
"value" }
            val out = lhs ?: rhs
            __check((out).toString(), "value")
            __check((order).toString(), "rhs")
        }
