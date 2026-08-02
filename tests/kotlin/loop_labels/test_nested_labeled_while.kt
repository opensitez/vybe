// vybe-test: kotlin/loop_labels/test_nested_labeled_while
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            var i = 0
            var out = 0
            outer@ while (i < 3) {
                var j = 0
                while (j < 3) {
                    if (i == 1 && j == 1) break@outer
                    out += i + j
                    j += 1
                }
                i += 1
            }
            println(out)
        }

