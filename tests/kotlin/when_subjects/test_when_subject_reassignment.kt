// vybe-test: kotlin/when_subjects/test_when_subject_reassignment
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var x = 1
            val out = when (x) {
                1 -> {
                    x = 2
                    "one"
                }
                else -> "other"
            }
            __check((out).toString(), "one")
            __check((x).toString(), "2")
        }
