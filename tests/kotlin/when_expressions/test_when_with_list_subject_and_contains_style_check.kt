// vybe-test: kotlin/when_expressions/test_when_with_list_subject_and_contains_style_check
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun classify(value: Int): String {
            return when (value) {
                in listOf(1, 3, 5) -> "odd-primeish"
                in listOf(2, 4, 6) -> "even-small"
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
            __check((classify(1)).toString(), "odd-primeish")
            __check((classify(4)).toString(), "even-small")
            __check((classify(7)).toString(), "other")
        }
