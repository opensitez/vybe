// vybe-test: kotlin/repeat_statements/test_repeat_long_count
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var out = 0L
            repeat(5) { i ->
                out += i.toLong()
            }
            __check((out).toString(), "10")
        }
