// vybe-test: kotlin/loop_labels/test_label_on_while_block
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var i = 0
            var out = 0
            loop@ while (true) {
                i += 1
                if (i == 5) break@loop
                out += i
            }
            println(out)
        }

