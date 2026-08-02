// vybe-test: kotlin/loops/test_do_while_runs_once_before_condition_is_checked
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var total = 0
            var i = 9
            do {
                total += i
                i -= 2
            } while (i < 0)
            println(total)
            println(i)
        }

