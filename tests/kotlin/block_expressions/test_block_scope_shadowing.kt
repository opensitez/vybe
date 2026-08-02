// vybe-test: kotlin/block_expressions/test_block_scope_shadowing
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = 1
val b = run { val a = 2
a + 1 }
__check((a + b).toString(), "4") }
