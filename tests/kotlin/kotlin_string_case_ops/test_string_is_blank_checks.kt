// vybe-test: kotlin/kotlin_string_case_ops/test_string_is_blank_checks
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_case_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("   ".isBlank()).toString(), "true")
            __check(("x".isBlank()).toString(), "false")
        }
