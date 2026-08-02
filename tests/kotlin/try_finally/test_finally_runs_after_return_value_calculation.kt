// vybe-test: kotlin/try_finally/test_finally_runs_after_return_value_calculation
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

fun f(): Int {
        try {
            return 1 + 1
        } finally {
            __check(("finally").toString(), "finally")
        }
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((f()).toString(), "2") }
