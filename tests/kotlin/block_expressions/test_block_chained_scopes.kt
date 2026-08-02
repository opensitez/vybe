// vybe-test: kotlin/block_expressions/test_block_chained_scopes
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val x = run {
            val a = 1
            run {
                val b = 2
                a + b
            }
        }
        __check((x).toString(), "3")
    }
