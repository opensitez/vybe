// vybe-test: kotlin/loop_labels/test_labelled_break_if_nested
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var out = 0
            outer@ for (i in 1..4) {
                for (j in 1..4) {
                    if (i == 3 && j == 2) break@outer
                    out += 1
                }
            }
            println(out)
        }

