// vybe-test: kotlin/when_expressions/test_when_with_subject_as_nullable_and_null_branch
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun label(value: String?): String {
            return when (value) {
                null -> "null"
                "" -> "empty"
                else -> "value"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label(null)).toString(), "null")
            __check((label("")).toString(), "empty")
            __check((label("ok")).toString(), "value")
        }
