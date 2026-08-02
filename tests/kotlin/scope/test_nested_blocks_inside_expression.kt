// vybe-test: kotlin/scope/test_nested_blocks_inside_expression
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 2
            val result = {
                val value = 5
                value * 3
            }
            __check((value).toString(), "2")
            __check((result).toString(), "15")
        }
