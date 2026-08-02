// vybe-test: kotlin/control_flow/test_while_loop_with_labeled_break
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var trace = ""
            var outer = 0
            outer@ while (outer < 4) {
                var inner = 0
                while (inner < 4) {
                    if (inner == 2) {
                        break@outer
                    }
                    trace += "${outer}-${inner};"
                    inner += 1
                }
                outer += 1
            }
            println(trace)
        }

