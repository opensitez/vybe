// vybe-test: kotlin/when_expressions/test_when_with_variable_subject_binding
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 7
            val label = when (value) {
                is Int -> "int-" + value
                else -> "none"
            }
            __check((label).toString(), "int-7")
        }
