// vybe-test: kotlin/labeled_control_flow/test_while_loop_label_break
// origin: languages/kotlin/tests/kotlin/test_labeled_control_flow.rs

fun main() {
            var i = 0
            var total = 0
            repeat@ while (i < 6) {
                i += 1
                if (i % 2 == 0) {
                    continue@repeat
                }
                if (i > 4) {
                    break@repeat
                }
                total += i
            }
            println(total)
        }

