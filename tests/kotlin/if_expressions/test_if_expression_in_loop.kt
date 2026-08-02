// vybe-test: kotlin/if_expressions/test_if_expression_in_loop
// origin: languages/kotlin/tests/kotlin/test_if_expressions.rs

fun main() {
            val values = listOf(1, 2, 3)
            var out = 0
            for (v in values) {
                out += if (v % 2 == 0) 2 else 1
            }
            println(out)
        }

