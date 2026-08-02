// vybe-test: kotlin/loop_labels/test_continue_label_with_for
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var out = 0
            outer@ for (i in 1..4) {
                for (j in 1..4) {
                    if (j == 1) continue@outer
                    out += i
                }
            }
            println(out)
        }

