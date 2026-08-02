// vybe-test: kotlin/do_while_control/test_do_while_continue_shape
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var i = 0
            var out = 0
            do {
                i += 1
                if (i == 2) continue
                out += i
            } while (i < 5)
            println(out)
        }

