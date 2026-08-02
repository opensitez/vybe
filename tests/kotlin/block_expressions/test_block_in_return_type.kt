// vybe-test: kotlin/block_expressions/test_block_in_return_type
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun f(v: Int): Int = if (v == 0) { 0 } else { val x = v * 2
x }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((f(4)).toString(), "8") }
