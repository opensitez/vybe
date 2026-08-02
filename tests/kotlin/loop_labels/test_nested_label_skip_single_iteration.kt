// vybe-test: kotlin/loop_labels/test_nested_label_skip_single_iteration
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var out = 0
            outer@ for (i in 1..4) {
                for (j in 1..4) {
                    if (j == 3) continue@outer
                    out += j
                }
            }
            println(out)
        }

