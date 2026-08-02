// vybe-test: kotlin/block_expressions/test_nested_blocks
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = { val b = { 2 + 1 }
b() + 1 }
__check((a()).toString(), "4") }
