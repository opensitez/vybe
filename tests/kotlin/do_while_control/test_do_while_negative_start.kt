// vybe-test: kotlin/do_while_control/test_do_while_negative_start
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var i = -3
            var out = 0
            do {
                out += i
                i += 1
            } while (i <= 0)
            println(out)
            println(i)
        }

