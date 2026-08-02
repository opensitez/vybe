// vybe-test: kotlin/repeat_statements/test_repeat_edge_zero
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var out = 0
            repeat(0) {
                out = 99
            }
            __check((out).toString(), "0")
        }
