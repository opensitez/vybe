// vybe-test: kotlin/repeat_statements/test_repeat_while_mixed
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var i = 0
            var out = 0
            repeat(5) {
                if (i % 2 == 0) out += i
                i += 1
            }
            __check((out).toString(), "6")
        }
