// vybe-test: kotlin/block_expressions/test_block_resulting_list
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x = run { listOf(1, 2, 3).map { it * 2 } }
__check((x[1]).toString(), "4") }
