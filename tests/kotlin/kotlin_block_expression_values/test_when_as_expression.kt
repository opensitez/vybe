// vybe-test: kotlin/kotlin_block_expression_values/test_when_as_expression
// origin: languages/kotlin/tests/kotlin/test_kotlin_block_expression_values.rs

fun whenResult(value: Int): String {
            return when (value) {
                1 -> "one"
                2 -> "two"
                else -> "many"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((whenResult(2)).toString(), "two")
            __check((whenResult(4)).toString(), "many")
        }
