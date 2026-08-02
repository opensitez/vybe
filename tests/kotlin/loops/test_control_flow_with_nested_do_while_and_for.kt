// vybe-test: kotlin/loops/test_control_flow_with_nested_do_while_and_for
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var rounds = 0
            var total = 0
            do {
                for (i in 1..3) {
                    total += i
                }
                rounds += 1
            } while (rounds < 2)
            println(total)
        }

