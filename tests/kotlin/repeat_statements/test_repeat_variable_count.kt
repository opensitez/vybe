// vybe-test: kotlin/repeat_statements/test_repeat_variable_count
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val n = 4
            var out = 0
            repeat(n) { i ->
                out += i + 1
            }
            __check((out).toString(), "10")
        }
