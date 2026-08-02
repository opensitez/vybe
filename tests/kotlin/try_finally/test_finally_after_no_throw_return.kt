// vybe-test: kotlin/try_finally/test_finally_after_no_throw_return
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun run(): Int {
        return try {
            7
        } finally {
            __check(("done").toString(), "done")
        }
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((run()).toString(), "7") }
