// vybe-test: kotlin/loop_labels/test_label_for_outer_then_inner
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var out = ""
            outer@ for (i in 1..3) {
                inner@ for (j in 1..3) {
                    if (i == 2) continue@outer
                    out += i.toString() + j.toString()
                }
            }
            println(out)
        }

