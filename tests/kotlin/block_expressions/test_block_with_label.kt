// vybe-test: kotlin/block_expressions/test_block_with_label
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val x = run {
            var i = 0
            if (i == 0) {
                i = 2
            }
            i
        }
        __check((x).toString(), "2")
    }
