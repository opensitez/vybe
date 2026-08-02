// vybe-test: kotlin/when_guards/test_when_guarded_multiple_vars
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 2
            val b = 4
            val out = when {
                a + b > 10 -> "big"
                a * b == 8 -> "match"
                else -> "other"
            }
            __check((out).toString(), "match")
        }
