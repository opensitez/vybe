// vybe-test: kotlin/local_returns/test_local_return_in_run_block_string
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "abc"
            val out = run {
                if (text.isEmpty()) return@run "empty"
                text
            }
            __check((out).toString(), "abc")
        }
