// vybe-test: kotlin/loops/test_while_labeled_continue_targets_outer_loop
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var i = 0
            var total = 0
            outer@ while (i < 3) {
                i += 1
                var j = 0
                while (j < 3) {
                    j += 1
                    if (j == 2) continue@outer
                    total += i * j
                }
                total += 10
            }
            println(i)
            println(total)
        }

