// vybe-test: kotlin/control_flow/test_repeat_zero_with_side_effect
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

var seen = 0
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            repeat(0) {
                seen += 1
            }
            __check((seen).toString(), "0")
        }
