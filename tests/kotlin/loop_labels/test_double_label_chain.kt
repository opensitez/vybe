// vybe-test: kotlin/loop_labels/test_double_label_chain
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var out = 0
            one@ for (i in 1..3) {
                two@ for (j in 1..3) {
                    if (i == 2 && j == 2) continue@one
                    out += i + j
                }
            }
            println(out)
        }

