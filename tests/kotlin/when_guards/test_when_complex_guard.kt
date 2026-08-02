// vybe-test: kotlin/when_guards/test_when_complex_guard
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun canRun(v: Int): Boolean = when {
            v < 0 -> false
            v == 0 -> false
            v % 2 == 1 -> false
            else -> true
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((canRun(2)).toString(), "true")
            __check((canRun(3)).toString(), "false")
        }
