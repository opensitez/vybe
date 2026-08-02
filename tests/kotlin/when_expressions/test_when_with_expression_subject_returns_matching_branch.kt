// vybe-test: kotlin/when_expressions/test_when_with_expression_subject_returns_matching_branch
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun score(label: Int): String {
            return when (label) {
                0 -> "zero"
                1 -> "one"
                2, 3 -> "small"
                else -> "many"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((score(0)).toString(), "zero")
            __check((score(2)).toString(), "small")
            __check((score(7)).toString(), "many")
        }
