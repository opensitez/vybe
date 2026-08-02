// vybe-test: kotlin/for_loop_variants/test_for_mutation_and_visibility
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var values = intArrayOf(1, 2, 3)
            for (i in values.indices) {
                values[i] = values[i] * 2
            }
            println(values[0] + values[1] + values[2])
        }

