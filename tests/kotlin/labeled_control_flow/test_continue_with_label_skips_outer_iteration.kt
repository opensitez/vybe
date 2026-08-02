// vybe-test: kotlin/labeled_control_flow/test_continue_with_label_skips_outer_iteration
// origin: languages/kotlin/tests/kotlin/test_labeled_control_flow.rs

fun main() {
            var events = 0
            outer@ for (i in 0..2) {
                for (j in 0..2) {
                    if (i == 1) {
                        continue@outer
                    }
                    events += 1
                }
            }
            println(events)
        }

