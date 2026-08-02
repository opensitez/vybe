// vybe-test: kotlin/loop_labels/test_labelled_continue_when_none
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var out = 0
            outer@ for (i in 1..4) {
                if (i == 3) continue@outer
                out += i
            }
            println(out)
        }

