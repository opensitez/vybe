// vybe-test: kotlin/ranges/test_unbounded_step_expression_range
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            val step = 2
            var total = 0
            for (value in 1..7 step step) {
                total += value
            }
            println(total)
        }

