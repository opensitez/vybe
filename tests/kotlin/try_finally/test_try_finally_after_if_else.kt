// vybe-test: kotlin/try_finally/test_try_finally_after_if_else
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun run(v: Int): Int {
        return if (v > 0) {
            try { v } finally { __check(("pos").toString(), "pos") }
        } else {
            try { -v } finally { __check(("neg").toString(), "1") }
        }
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((run(1)).toString(), "neg")
__check((run(-2)).toString(), "2") }
