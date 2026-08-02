// vybe-test: kotlin/block_expressions/test_block_with_local_return
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val x = run {
            __check(("inside").toString(), "inside")
            5
        }
        __check((x).toString(), "5")
    }
