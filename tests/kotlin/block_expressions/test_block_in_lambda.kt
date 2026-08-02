// vybe-test: kotlin/block_expressions/test_block_in_lambda
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val fn = { x: Int ->
            run {
                val a = x + 1
                a * 2
            }
        }
        __check((fn(3)).toString(), "8")
    }
