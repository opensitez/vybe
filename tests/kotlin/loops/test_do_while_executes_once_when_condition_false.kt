// vybe-test: kotlin/loops/test_do_while_executes_once_when_condition_false
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var total = 0
            var i = 0
            do {
                total += i
                i += 1
            } while (false)
            println(total)
            println(i)
        }

