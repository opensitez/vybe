// vybe-test: kotlin/control_flow/test_for_loop_accumulation
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var total = 0
            for (i in 1..4) {
                total += i
            }
            println(total)
        }

