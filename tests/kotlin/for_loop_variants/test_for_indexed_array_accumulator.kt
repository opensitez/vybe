// vybe-test: kotlin/for_loop_variants/test_for_indexed_array_accumulator
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            val values = intArrayOf(2, 4, 6, 8)
            var out = 0
            for (i in values.indices) {
                if (i % 2 == 0) out += values[i]
            }
            println(out)
        }

