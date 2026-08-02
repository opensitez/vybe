// vybe-test: kotlin/when_guards/test_when_reusable_expression
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 7
            val out = when {
                x > 10 -> "gt"
                x % 2 == 0 -> "even"
                else -> "odd"
            }
            __check((out).toString(), "odd")
        }
