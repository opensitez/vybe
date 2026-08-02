// vybe-test: kotlin/for_loop_variants/test_for_array_elements
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            val values = intArrayOf(1, 2, 3)
            var total = 0
            for (v in values) {
                total += v
            }
            println(total)
        }

