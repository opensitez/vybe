// vybe-test: kotlin/when_subjects/test_when_subject_local_let
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 10
            val out = when (x) {
                5 -> 1
                10 -> {
                    val y = x / 2
                    y
                }
                else -> 0
            }
            __check((out).toString(), "5")
        }
