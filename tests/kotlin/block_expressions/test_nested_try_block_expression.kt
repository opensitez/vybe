// vybe-test: kotlin/block_expressions/test_nested_try_block_expression
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val out = run {
            try {
                1
            } catch (e: Exception) {
                0
            }
        }
        __check((out).toString(), "1")
    }
