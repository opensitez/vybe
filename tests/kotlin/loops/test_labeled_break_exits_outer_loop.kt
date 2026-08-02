// vybe-test: kotlin/loops/test_labeled_break_exits_outer_loop
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var total = 0
            outer@ for (i in 1..5) {
                for (j in 1..5) {
                    total += i + j
                    if (i == 3 && j == 2) break@outer
                }
            }
            println(total)
        }

