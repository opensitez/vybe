// vybe-test: kotlin/labeled_control_flow/test_break_inner_loop_does_not_leave_outer_loop
// origin: languages/kotlin/tests/kotlin/test_labeled_control_flow.rs

fun main() {
            var acc = 0
            outer@ for (i in 0..2) {
                for (j in 0..2) {
                    if (j == 1) {
                        break
                    }
                    acc += 1
                }
                acc += 10
            }
            println(acc)
        }

