// vybe-test: kotlin/control_flow/test_single_line_if
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 10
            if (x > 0) __check(("positive").toString(), "positive")
        }
