// vybe-test: kotlin/labeled_control_flow/test_nested_continue_labeling_controls_flow
// origin: languages/kotlin/tests/kotlin/test_labeled_control_flow.rs

fun main() {
            var values = 0
            outer@ for (i in 0..2) {
                for (j in 0..2) {
                    if (i + j < 2) {
                        continue@outer
                    }
                    values += 1
                }
            }
            println(values)
        }

