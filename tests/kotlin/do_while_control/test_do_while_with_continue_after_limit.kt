// vybe-test: kotlin/do_while_control/test_do_while_with_continue_after_limit
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var i = 0
            var total = 0
            do {
                i += 1
                if (i == 3) continue
                total += i
            } while (i < 6)
            println(total)
        }

