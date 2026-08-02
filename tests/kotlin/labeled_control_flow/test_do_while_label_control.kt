// vybe-test: kotlin/labeled_control_flow/test_do_while_label_control
// origin: languages/kotlin/tests/kotlin/test_labeled_control_flow.rs

fun main() {
            var n = 0
            var acc = 0
            run@ do {
                n += 1
                if (n == 3) {
                    continue@run
                }
                acc += n
            } while (n < 5)
            println(acc)
        }

