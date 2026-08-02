// vybe-test: kotlin/repeat_statements/test_repeat_after_zero_like
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var out = 1
            repeat(0) {
                out = 9
            }
            repeat(2) {
                out += out
            }
            __check((out).toString(), "4")
        }
