// vybe-test: kotlin/do_while_control/test_do_while_long_progression
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var i = 1L
            var out = 0L
            do {
                out += i
                i += 2
            } while (i < 8)
            println(out)
        }

