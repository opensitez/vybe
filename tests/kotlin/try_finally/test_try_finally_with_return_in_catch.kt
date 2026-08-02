// vybe-test: kotlin/try_finally/test_try_finally_with_return_in_catch
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun run(): String {
        try {
            throw RuntimeException("x")
        } catch (e: Exception) {
            return "err"
        } finally {
            __check(("fin").toString(), "fin")
        }
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((run()).toString(), "err") }
