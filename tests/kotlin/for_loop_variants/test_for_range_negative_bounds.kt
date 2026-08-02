// vybe-test: kotlin/for_loop_variants/test_for_range_negative_bounds
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var total = 0
            for (i in -2..2) {
                total += i
            }
            println(total)
        }

