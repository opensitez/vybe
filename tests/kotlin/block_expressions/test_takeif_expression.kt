// vybe-test: kotlin/block_expressions/test_takeif_expression
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x = 3
val y = x.takeIf { it > 2 }
__check((y).toString(), "3") }
