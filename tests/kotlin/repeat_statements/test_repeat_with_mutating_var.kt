// vybe-test: kotlin/repeat_statements/test_repeat_with_mutating_var
// origin: languages/kotlin/tests/kotlin/test_repeat_statements.rs

class C {
            var n = 0
            fun inc() { n += 1 }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = C()
            repeat(3) { c.inc() }
            __check((c.n).toString(), "3")
        }
