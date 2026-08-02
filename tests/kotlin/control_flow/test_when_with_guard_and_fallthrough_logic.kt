// vybe-test: kotlin/control_flow/test_when_with_guard_and_fallthrough_logic
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val score = 92
            val label = when (score) {
                in 90..100 -> if (score % 2 == 0) "A" else "A+"
                in 80..89 -> "B"
                else -> "F"
            }
            __check((label).toString(), "A")
        }
