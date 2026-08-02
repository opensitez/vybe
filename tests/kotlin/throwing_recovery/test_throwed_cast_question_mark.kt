// vybe-test: kotlin/throwing_recovery/test_throwed_cast_question_mark
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = "text"
            __check((value as? Int).toString(), "null")
            __check((value is Int).toString(), "false")
        }
