// vybe-test: kotlin/control_flow/test_when_without_subject_uses_guard_chain
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val score = 58
            val band = when {
                score >= 90 -> "A"
                score >= 80 -> "B"
                score >= 70 -> "C"
                score >= 60 -> "D"
                else -> "F"
            }
            __check((band).toString(), "D")
        }
