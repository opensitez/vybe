// vybe-test: kotlin/block_expressions/test_when_block_result
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val out = when (2) { 1 -> 10
2 -> run { val a = 5
a + 1 }
else -> 0 }
__check((out).toString(), "6") }
