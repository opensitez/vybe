// vybe-test: kotlin/repeat_statements/test_repeat_and_boolean
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var out = true
            repeat(2) {
                out = out && (it < 2)
            }
            __check((out).toString(), "true")
        }
