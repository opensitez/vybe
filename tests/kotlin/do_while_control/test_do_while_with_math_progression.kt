// vybe-test: kotlin/do_while_control/test_do_while_with_math_progression
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var i = 0
            var out = 0
            do {
                out += i * 2
                i += 1
            } while (i < 4)
            println(out)
        }

