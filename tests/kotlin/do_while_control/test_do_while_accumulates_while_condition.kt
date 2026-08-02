// vybe-test: kotlin/do_while_control/test_do_while_accumulates_while_condition
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var i = 0
            var out = 0
            do {
                out += i
                i += 1
            } while (i < 4)
            println(out)
            println(i)
        }

