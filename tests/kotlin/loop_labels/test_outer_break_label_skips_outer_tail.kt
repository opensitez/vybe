// vybe-test: kotlin/loop_labels/test_outer_break_label_skips_outer_tail
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var out = 0
            outer@ for (i in 1..4) {
                for (j in 1..4) {
                    if (i == 2) break@outer
                    out += i + j
                }
            }
            println(out)
        }

