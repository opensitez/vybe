// vybe-test: kotlin/loops/test_do_while_executes_until_condition_fails
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var total = 0
            var i = 1
            do {
                total += i
                i += 1
            } while (i <= 4)
            println(total)
        }

