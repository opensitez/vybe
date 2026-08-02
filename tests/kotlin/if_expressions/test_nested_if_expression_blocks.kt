// vybe-test: kotlin/if_expressions/test_nested_if_expression_blocks
// origin: languages/kotlin/tests/kotlin/test_if_expressions.rs

fun label(v: Int): String {
            return if (v % 2 == 0) {
                if (v > 10) "large-even" else "small-even"
            } else {
                if (v > 10) "large-odd" else "small-odd"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label(4)).toString(), "small-even")
            __check((label(9)).toString(), "small-odd")
            __check((label(12)).toString(), "large-even")
        }
