// vybe-test: kotlin/repeat_statements/test_repeat_boolean_flags
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var out = true
            repeat(3) {
                out = out && it != 1
            }
            __check((out).toString(), "false")
        }
