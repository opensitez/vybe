// vybe-test: kotlin/for_loop_variants/test_for_array_with_index
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            val values = intArrayOf(4, 5, 6)
            var out = 0
            for (i in values.indices) {
                out += values[i]
            }
            println(out)
        }

