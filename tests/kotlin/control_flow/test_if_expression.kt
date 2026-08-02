// vybe-test: kotlin/control_flow/test_if_expression
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 7
            val b = 12
            val max = if (a > b) a else b
            __check((max).toString(), "12")
        }
