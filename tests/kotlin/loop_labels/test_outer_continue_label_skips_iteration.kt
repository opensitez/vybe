// vybe-test: kotlin/loop_labels/test_outer_continue_label_skips_iteration
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var out = 0
            outer@ for (i in 1..3) {
                for (j in 1..2) {
                    if (j == 2) continue@outer
                    out += i
                }
            }
            println(out)
        }

