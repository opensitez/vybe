// vybe-test: kotlin/block_expressions/test_block_expression_in_map
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x = mapOf(1 to run { 2 + 3 }, 2 to run { 5 + 6})
__check((x[1]!! + x[2]!!).toString(), "16") }
