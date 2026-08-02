// vybe-test: kotlin/loop_labels/test_label_nested_label_names
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var out = 0
            one@ for (i in 1..3) {
                two@ for (j in 1..3) {
                    three@ for (k in 1..3) {
                        if (i == 2 && j == 2 && k == 2) break@two
                        out += 1
                    }
                }
            }
            println(out)
        }

