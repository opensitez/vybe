// vybe-test: kotlin/control_flow/test_if_expression_used_in_assignment
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val count = 6
            val label = if (count > 10) "large" else if (count >= 5) "medium" else "small"
            __check((label).toString(), "medium")
        }
