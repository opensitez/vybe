// vybe-test: kotlin/block_expressions/test_block_with_return_from_inner_function
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val x = run {
            fun y() = 2
            y() + 3
        }
        __check((x).toString(), "5")
    }
