// vybe-test: kotlin/try_finally/test_try_finally_with_local_return
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun run(): String {
        var v = ""
        val result = run {
            try {
                "try"
            } finally {
                v = "finally"
            }
        }
        return v + ":" + result
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((run()).toString(), "finally:try") }
