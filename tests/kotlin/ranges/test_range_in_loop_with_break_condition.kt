// vybe-test: kotlin/ranges/test_range_in_loop_with_break_condition
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            var total = 0
            for (value in 1..20) {
                if (value == 7) {
                    break
                }
                total += value
            }
            println(total)
        }

