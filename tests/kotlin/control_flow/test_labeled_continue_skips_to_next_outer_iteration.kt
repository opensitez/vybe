// vybe-test: kotlin/control_flow/test_labeled_continue_skips_to_next_outer_iteration
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var values = ""
            outer@ for (outer in 1..3) {
                inner@ for (inner in 1..3) {
                    if (inner == 2) continue@outer
                    values += "${outer}${inner};"
                }
            }
            println(values)
        }

