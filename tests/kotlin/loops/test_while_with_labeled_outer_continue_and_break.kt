// vybe-test: kotlin/loops/test_while_with_labeled_outer_continue_and_break
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var i = 0
            var sum = 0
            outer@ while (i < 10) {
                i += 1
                if (i == 3) continue@outer
                if (i == 8) break@outer
                sum += i
            }
            println(sum)
        }

