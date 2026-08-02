// vybe-test: kotlin/block_expressions/test_block_in_sequence
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x = sequenceOf(1, 2, 3).map { it * run { 2 } }.sum()
__check((x).toString(), "12") }
