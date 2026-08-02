// vybe-test: kotlin/loop_labels/test_do_while_with_label
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var i = 0
            var out = 0
            mark@ do {
                out += i
                i += 1
            } while (i < 3)
            println(out)
        }

