// vybe-test: kotlin/kotlin_block_expression_values/test_if_as_expression
// origin: languages/kotlin/tests/kotlin/test_kotlin_block_expression_values.rs

fun classify(value: Int): String {
            return if (value > 0) {
                "pos"
            } else {
                "non"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(1)).toString(), "pos")
            __check((classify(0)).toString(), "non")
        }
