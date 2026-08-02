// vybe-test: kotlin/block_expressions/test_block_as_expression_in_variable
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x = if (true) { 1 + 2 } else { 3 }
__check((x).toString(), "3") }
