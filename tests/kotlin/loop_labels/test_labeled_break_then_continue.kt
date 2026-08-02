// vybe-test: kotlin/loop_labels/test_labeled_break_then_continue
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var out = 0
            for (i in 1..4) {
                block@ for (j in 1..4) {
                    if (j == 2) break@block
                    if (i == 3) continue
                    out += 1
                }
            }
            println(out)
        }

