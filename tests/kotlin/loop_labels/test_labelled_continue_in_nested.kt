// vybe-test: kotlin/loop_labels/test_labelled_continue_in_nested
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var out = 0
            outer@ for (i in 1..3) {
                for (j in 1..3) {
                    if (i == 2) continue@outer
                    out += j
                }
            }
            println(out)
        }

