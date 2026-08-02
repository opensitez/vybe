// vybe-test: kotlin/when_subjects/test_when_subject_reassigning
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var value = 4
            val out = when (value) {
                4 -> {
                    value += 1
                    "four"
                }
                else -> "other"
            }
            __check((out).toString(), "four")
            __check((value).toString(), "5")
        }
