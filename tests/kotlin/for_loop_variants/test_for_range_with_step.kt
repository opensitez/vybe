// vybe-test: kotlin/for_loop_variants/test_for_range_with_step
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var total = 0
            for (i in 1..10 step 3) {
                total += i
            }
            println(total)
        }

