// vybe-test: kotlin/repeat_statements/test_repeat_mutable_state
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

var acc = 0
        fun bump() {
            acc += 1
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            repeat(7) { bump() }
            __check((acc).toString(), "7")
        }
