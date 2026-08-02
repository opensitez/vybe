// vybe-test: kotlin/block_expressions/test_block_resulting_map
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x = run { val a = mapOf(1 to 2, 3 to 4)
a }
__check((x[3]).toString(), "4") }
