// vybe-test: kotlin/when_guards/test_when_guarded_array_bounds
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 3
            val out = when {
                value in 1..3 -> "small"
                value in 4..6 -> "med"
                else -> "large"
            }
            __check((out).toString(), "small")
        }
