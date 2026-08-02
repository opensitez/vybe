// vybe-test: kotlin/loop_labels/test_label_on_while_with_continue
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var i = 0
            var out = 0
            outer@ while (i < 6) {
                i += 1
                if (i % 2 == 0) continue@outer
                out += i
            }
            println(out)
        }

