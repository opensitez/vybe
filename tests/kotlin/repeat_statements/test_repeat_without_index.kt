// vybe-test: kotlin/repeat_statements/test_repeat_without_index
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var out = 0
            repeat(4) {
                out = out * 10 + 1
            }
            __check((out).toString(), "1111")
        }
