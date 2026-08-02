// vybe-test: kotlin/block_expressions/test_with_expression
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val out = with(StringBuilder()) { append("a")
append("b")
toString() }
__check((out).toString(), "ab") }
