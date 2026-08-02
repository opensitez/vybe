// vybe-test: kotlin/when_guards/test_when_with_nested_guard_scope
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val threshold = 4
            val value = 3
            val out = when {
                value > threshold -> "too-high"
                else -> {
                    val scaled = value + threshold
                    if (scaled > 5) "scaled" else "small"
                }
            }
            __check((out).toString(), "scaled")
        }
