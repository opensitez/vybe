// vybe-test: kotlin/while_loops/test_while_true_with_break_condition
// origin: languages/kotlin/tests/kotlin/test_while_loops.rs

fun main() {
            var i = 0
            var total = 0
            while (true) {
                if (i == 4) break
                total += i
                i += 1
            }
            println(total)
        }

