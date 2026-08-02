// vybe-test: kotlin/control_flow/test_labeled_break_exits_outer_loop_only
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var values = ""
            outer@ for (outer in 1..3) {
                for (inner in 1..3) {
                    if (outer == 2 && inner == 2) {
                        break@outer
                    }
                    values += "${outer}${inner};"
                }
            }
            println(values)
        }

