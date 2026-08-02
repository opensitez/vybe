// vybe-test: kotlin/control_flow/test_repeat_zero_iteration
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var count = 0
            repeat(0) { count += 1 }
            __check((count).toString(), "0")
        }
