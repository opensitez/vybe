// vybe-test: kotlin/block_expressions/test_block_with_multiple_statements
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x = run { __check(("a").toString(), "a")
__check(("b").toString(), "b")
3 }
__check((x).toString(), "3") }
