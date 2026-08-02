// vybe-test: kotlin/loop_labels/test_label_chain_in_nested_for
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var out = 0
            alpha@ for (i in 1..3) {
                beta@ for (j in 1..3) {
                    if (i + j == 6) break@alpha
                    out += 1
                }
            }
            println(out)
        }

