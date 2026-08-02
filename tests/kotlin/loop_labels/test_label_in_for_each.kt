// vybe-test: kotlin/loop_labels/test_label_in_for_each
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var out = ""
            outer@ for (x in intArrayOf(1,2,3)) {
                for (y in intArrayOf(4,5)) {
                    if (y == 5) continue@outer
                    out += y.toString()
                }
            }
            println(out)
        }

