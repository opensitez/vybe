// vybe-test: kotlin/ranges/test_range_in_loop_with_continue_condition
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            var total = 0
            for (value in 1..10) {
                if (value % 2 == 0) {
                    continue
                }
                total += value
            }
            println(total)
        }

