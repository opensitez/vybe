// vybe-test: kotlin/when_expressions/test_when_with_inclusive_ranges_and_open_edges
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun level(value: Int): String {
            return when (value) {
                Int.MIN_VALUE..-1 -> "negative"
                0 -> "zero"
                1..99 -> "low"
                100..Int.MAX_VALUE -> "high"
                else -> "other"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((level(-4)).toString(), "negative")
            __check((level(0)).toString(), "zero")
            __check((level(1)).toString(), "low")
            __check((level(100)).toString(), "high")
        }
