// vybe-test: kotlin/when_expressions/test_when_with_multiple_subject_expressions
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun classify(value: Int): String {
            return when (value) {
                in 1..3 -> "low"
                in 4..10 -> "mid"
                !in 1..10 -> "outside"
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
            __check((classify(2)).toString(), "low")
            __check((classify(10)).toString(), "mid")
            __check((classify(20)).toString(), "outside")
        }
