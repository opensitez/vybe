// vybe-test: kotlin/do_while_control/test_do_while_with_zero_iterations_guard
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var i = 0
            var out = 0
            do {
                out += 1
                i += 1
            } while (false)
            println(out)
            println(i)
        }

