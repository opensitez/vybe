// vybe-test: kotlin/when_subjects/test_when_subject_with_boolean
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = true
            val out = when (x) {
                true -> "ok"
                false -> "no"
            }
            __check((out).toString(), "ok")
        }
