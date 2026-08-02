// vybe-test: kotlin/block_expressions/test_block_resulting_string
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x = run { val a = "x"
val b = "y"
a + b }
__check((x).toString(), "xy") }
