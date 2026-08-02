// vybe-test: kotlin/loop_labels/test_repeated_labeled_for
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var out = 0
            for (i in 1..5) {
                mark@ for (j in 1..5) {
                    if (i + j > 6) break@mark
                    out += 1
                }
            }
            println(out)
        }

