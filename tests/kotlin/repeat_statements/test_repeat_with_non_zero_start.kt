// vybe-test: kotlin/repeat_statements/test_repeat_with_non_zero_start
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var i = 3
            var out = 0
            repeat(4) {
                out += i
                i += 1
            }
            __check((out).toString(), "22")
        }
