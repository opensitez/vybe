// vybe-test: kotlin/labeled_control_flow/test_break_with_loop_label_stops_outer_loop
// origin: languages/kotlin/tests/kotlin/test_labeled_control_flow.rs

fun main() {
            var total = 0
            outer@ for (i in 0..4) {
                if (i == 4) {
                    break@outer
                }
                if (i == 2) {
                    continue
                }
                total += i
            }
            println(total)
        }

