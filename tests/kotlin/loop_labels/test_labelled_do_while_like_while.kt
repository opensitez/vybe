// vybe-test: kotlin/loop_labels/test_labelled_do_while_like_while
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var i = 0
            var out = 0
            block@ while (i < 5) {
                out += i
                i += 1
                continue@block
            }
            println(out)
        }

