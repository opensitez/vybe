// vybe-test: kotlin/strings/test_string_repeat_when_count_zero
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("x".repeat(0)).toString(), "")
            __check(("x".padStart(0)).toString(), "x")
            __check(("x".padEnd(0)).toString(), "x")
        }
