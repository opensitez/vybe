// vybe-test: kotlin/for_loop_variants/test_for_nested_range_sum
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var total = 0
            for (i in 1..3) {
                for (j in 1..2) {
                    total += i * j
                }
            }
            println(total)
        }

