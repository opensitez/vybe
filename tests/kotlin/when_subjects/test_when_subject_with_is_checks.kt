// vybe-test: kotlin/when_subjects/test_when_subject_with_is_checks
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun describe(x: Any): String = when (x) {
            is Int -> "int"
            is Double -> "double"
            is String -> "string"
            else -> "other"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe("x")).toString(), "string")
            __check((describe(4.5)).toString(), "double")
        }
