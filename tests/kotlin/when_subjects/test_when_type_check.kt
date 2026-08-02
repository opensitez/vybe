// vybe-test: kotlin/when_subjects/test_when_type_check
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun tag(v: Any): String = when (v) {
            is Int -> "int"
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
            __check((tag(5)).toString(), "int")
            __check((tag("x")).toString(), "string")
            __check((tag(1.5)).toString(), "other")
        }
