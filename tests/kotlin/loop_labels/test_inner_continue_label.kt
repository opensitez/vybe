// vybe-test: kotlin/loop_labels/test_inner_continue_label
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var out = 0
            for (i in 1..3) {
                inner@ for (j in 1..3) {
                    if (j == 2) continue@inner
                    out += j
                }
            }
            println(out)
        }

