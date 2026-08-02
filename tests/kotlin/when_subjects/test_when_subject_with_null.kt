// vybe-test: kotlin/when_subjects/test_when_subject_with_null
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun safeDescribe(v: Int?): String = when (v) {
            null -> "null"
            0 -> "zero"
            else -> "other"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((safeDescribe(null)).toString(), "null")
            __check((safeDescribe(0)).toString(), "zero")
        }
