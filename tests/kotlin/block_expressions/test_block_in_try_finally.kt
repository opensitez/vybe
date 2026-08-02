// vybe-test: kotlin/block_expressions/test_block_in_try_finally
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val x = try {
            run {
                1 + 2
            }
        } finally {
            __check(("done").toString(), "done")
        }
        __check((x).toString(), "3")
    }
