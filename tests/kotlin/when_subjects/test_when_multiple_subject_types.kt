// vybe-test: kotlin/when_subjects/test_when_multiple_subject_types
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun classify(v: Any): String = when (v) {
            is Int -> "number"
            is Char -> "char"
            is Boolean -> "bool"
            else -> "other"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify('a')).toString(), "char")
            __check((classify(false)).toString(), "bool")
        }
