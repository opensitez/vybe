// vybe-test: kotlin/when_guards/test_when_guarded_short_circuit
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 5
            val y = 0
            val out = when {
                x > 0 && y == 0 -> "safe"
                x > 10 && y == 1 -> "skip"
                else -> "other"
            }
            __check((out).toString(), "safe")
        }
