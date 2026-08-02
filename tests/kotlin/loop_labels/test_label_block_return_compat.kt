// vybe-test: kotlin/loop_labels/test_label_block_return_compat
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var out = 0
            one@ for (i in 1..2) {
                two@ for (j in 1..2) {
                    if (i == 2 && j == 2) break@one
                    out += j
                }
            }
            println(out)
        }

