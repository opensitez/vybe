// vybe-test: kotlin/control_flow/test_while_loop_with_labeled_continue
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var trace = ""
            var rounds = 0
            outer@ while (rounds < 3) {
                rounds += 1
                var inner = 0
                while (inner < 4) {
                    inner += 1
                    if (inner == 2) {
                        continue@outer
                    }
                    trace += inner.toString()
                }
                trace += "x"
            }
            println(trace)
        }

