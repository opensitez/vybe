// vybe-test: kotlin/control_flow/test_nested_if_expressions
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 5
            val b = 10
            val c = 15
            val max = if (a > b) (if (a > c) a else c) else (if (b > c) b else c)
            __check((max).toString(), "15")
        }
