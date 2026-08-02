// vybe-test: kotlin/loops/test_labeled_continue_skips_outer_iteration
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var total = 0
            outer@ for (i in 1..4) {
                for (j in 1..3) {
                    if (j == 3 && i == 2) continue@outer
                    total += i * j
                }
            }
            println(total)
        }

