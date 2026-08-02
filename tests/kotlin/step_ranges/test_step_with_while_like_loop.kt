// vybe-test: kotlin/step_ranges/test_step_with_while_like_loop
// origin: languages/kotlin/tests/kotlin/test_step_ranges.rs

fun main() {
            var count = 0
            for (v in 0..10 step 2) {
                count += v
            }
            println(count)
        }

