// vybe-test: kotlin/block_expressions/test_block_for_mutation_then_use
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x = run { var a = 1
a += 2
a * 2 }
__check((x).toString(), "6") }
