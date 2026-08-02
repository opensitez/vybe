// vybe-test: kotlin/evaluation_order/test_elvis_skips_rhs_when_non_null
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val lhs: String? = "left"
            var order = ""
            val out = lhs ?: run { order += "rhs"
"value" }
            __check((out).toString(), "left")
            __check((order).toString(), "")
        }
