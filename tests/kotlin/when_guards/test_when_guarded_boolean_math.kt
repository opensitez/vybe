// vybe-test: kotlin/when_guards/test_when_guarded_boolean_math
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = when {
                1 + 1 == 3 -> "wrong"
                1 + 1 == 2 -> "yes"
                else -> "no"
            }
            __check((out).toString(), "yes")
        }
