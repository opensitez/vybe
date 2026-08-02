// vybe-test: kotlin/block_expressions/test_block_with_conditional_cast
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val v: Any = 3
val x = if (v is Int) { v + 1 } else { -1 }
__check((x).toString(), "4") }
