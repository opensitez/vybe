// vybe-test: kotlin/for_loop_variants/test_for_range_long_steps
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

fun main() {
            var count = 0L
            for (i in 1L..7L step 2) {
                count += i
            }
            println(count)
        }

