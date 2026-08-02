// vybe-test: kotlin/control_flow/test_for_range_with_step
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var sum = 0
            for (i in 1..7 step 2) {
                sum += i
            }
            println(sum)
        }

