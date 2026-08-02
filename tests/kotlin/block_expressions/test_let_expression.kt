// vybe-test: kotlin/block_expressions/test_let_expression
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x = "a".let { it + "b" }
__check((x).toString(), "ab") }
