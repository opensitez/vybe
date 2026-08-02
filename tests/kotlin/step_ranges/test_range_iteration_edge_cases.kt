// vybe-test: kotlin/step_ranges/test_range_iteration_edge_cases
// origin: languages/kotlin/tests/kotlin/test_step_ranges.rs

fun main() {
            var total = 0
            for (x in 0 until 1) {
                total += x
            }
            var done = 0
            for (x in 2 downTo 2) {
                done += x
            }
            println(total)
            println(done)
        }

