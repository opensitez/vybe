// vybe-test: kotlin/loops/test_for_in_range_inside_while_interaction
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var rounds = 0
            while (rounds < 2) {
                rounds += 1
                var row = 0
                for (i in 1..3) {
                    row += i
                }
                println(row)
            }
        }

