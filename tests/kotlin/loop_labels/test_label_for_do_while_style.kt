// vybe-test: kotlin/loop_labels/test_label_for_do_while_style
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var i = 0
            var out = ""
            mark@ for (ch in 1..3) {
                out += ch.toString()
                if (ch == 2) continue@mark
                out += "x"
            }
            println(out)
        }

