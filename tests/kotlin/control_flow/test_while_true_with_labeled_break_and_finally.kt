// vybe-test: kotlin/control_flow/test_while_true_with_labeled_break_and_finally
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var steps = 0
            try {
                outer@ while (true) {
                    steps += 1
                    if (steps == 2) {
                        break@outer
                    }
                }
            } finally {
                println(steps)
            }
            println("done")
        }

