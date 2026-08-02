// vybe-test: kotlin/literals/test_long_integer_literals
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val short = 10L
            val big = 1_000_000_000L
            val neg = -12L
            __check((short).toString(), "10")
            __check((big).toString(), "1000000000")
            __check((neg).toString(), "-12")
            __check((short + neg).toString(), "-2")
        }
