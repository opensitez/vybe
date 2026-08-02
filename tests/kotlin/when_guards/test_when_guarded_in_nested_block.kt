// vybe-test: kotlin/when_guards/test_when_guarded_in_nested_block
// origin: languages/kotlin/tests/kotlin/test_when_guards.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 7
            val out = when {
                value < 0 -> "neg"
                value < 5 -> "low"
                else -> {
                    if (value % 2 == 1) "odd" else "even"
                }
            }
            __check((out).toString(), "odd")
        }
