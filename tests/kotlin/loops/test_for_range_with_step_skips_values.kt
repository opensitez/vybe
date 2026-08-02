// vybe-test: kotlin/loops/test_for_range_with_step_skips_values
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var output = 0
            for (i in 1..10 step 3) {
                output += i
            }
            println(output)
        }

