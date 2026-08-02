// vybe-test: kotlin/for_loop_variants/test_for_for_expression_output
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            val rows = intArrayOf(1, 2)
            val cols = intArrayOf(10, 20)
            var out = 0
            for (r in rows) {
                for (c in cols) {
                    out += r + c
                }
            }
            println(out)
        }

