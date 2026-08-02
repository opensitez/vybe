// vybe-test: kotlin/when_guards/test_when_with_boolean_guard
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun isReady(v: Int): Boolean = v > 1 && v < 5
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = when {
                isReady(0) -> "no"
                isReady(3) -> "yes"
                else -> "nope"
            }
            __check((out).toString(), "yes")
        }
