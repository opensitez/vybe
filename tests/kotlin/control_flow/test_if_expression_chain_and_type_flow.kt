// vybe-test: kotlin/control_flow/test_if_expression_chain_and_type_flow
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 10
            val tier = if (x > 20) "high" else if (x > 5) "mid" else "low"
            if (x > 5) {
                __check((tier).toString(), "mid")
                __check(("statement").toString(), "statement")
            }
        }
