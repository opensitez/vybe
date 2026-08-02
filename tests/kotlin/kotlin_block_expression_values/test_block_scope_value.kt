// vybe-test: kotlin/kotlin_block_expression_values/test_block_scope_value
// origin: languages/kotlin/tests/kotlin/test_kotlin_block_expression_values.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = {
                val x = 2
                val y = 3
                x + y
            }
            __check((result).toString(), "5")
        }
