// vybe-test: kotlin/do_while_control/test_do_while_large_jump
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var i = 0
            var out = 0
            do {
                out += i
                i = 100
            } while (i < 10)
            println(out)
        }

