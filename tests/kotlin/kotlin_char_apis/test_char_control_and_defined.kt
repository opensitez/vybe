// vybe-test: kotlin/kotlin_char_apis/test_char_control_and_defined
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('\u0000'.isISOControl()).toString(), "true")
            __check(('A'.isISOControl()).toString(), "false")
            __check(('A'.isDefined()).toString(), "true")
            __check(('\uFFFF'.isDefined()).toString(), "false")
        }
