// vybe-test: kotlin/while_loops/test_while_loop_basic_sum
// origin: languages/kotlin/tests/kotlin/test_while_loops.rs

fun main() {
            var i = 0
            var total = 0
            while (i < 5) {
                total += i
                i += 1
            }
            println(total)
        }

