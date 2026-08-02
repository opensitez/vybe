// vybe-test: kotlin/when_subjects/test_when_inline_subject
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val n = 8
            val result = when (n) {
                1 -> "one"
                8 -> "eight"
                else -> "other"
            }
            __check((result).toString(), "eight")
        }
