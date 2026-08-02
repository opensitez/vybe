// vybe-test: kotlin/loop_labels/test_label_without_jump
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var out = 0
            outer@ for (i in 1..2) {
                out += i
            }
            println(out)
        }

