// vybe-test: kotlin/loop_labels/test_label_with_while_and_if
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var i = 0
            var out = 0
            outer@ while (i < 6) {
                i += 1
                if (i == 4) continue@outer
                if (i == 5) break@outer
                out += i
            }
            println(out)
        }

