// vybe-test: kotlin/repeat_statements/test_repeat_in_tail_function
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

fun inc(n: Int): Int {
            var out = 0
            repeat(n) { out += 1 }
            return out
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((inc(5)).toString(), "5")
        }
