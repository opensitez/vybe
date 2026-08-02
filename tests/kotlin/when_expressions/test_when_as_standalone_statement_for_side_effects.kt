// vybe-test: kotlin/when_expressions/test_when_as_standalone_statement_for_side_effects
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var acc = ""
            val value = 4
            when {
                value > 10 -> acc = "big"
                value > 1 -> acc = "mid"
                else -> acc = "small"
            }
            __check((acc).toString(), "mid")
        }
