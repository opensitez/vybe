// vybe-test: kotlin/when_subjects/test_when_subject_default_for_unknown
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun status(x: String): String = when (x) {
            "on" -> "1"
            "off" -> "0"
            else -> "x"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((status("on")).toString(), "1")
            __check((status("pause")).toString(), "x")
        }
