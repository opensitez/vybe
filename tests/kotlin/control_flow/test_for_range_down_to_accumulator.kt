// vybe-test: kotlin/control_flow/test_for_range_down_to_accumulator
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var total = 0
            for (i in 8 downTo 3) {
                total += i
            }
            println(total)
        }

