// vybe-test: kotlin/do_while_control/test_do_while_with_break
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var i = 0
            var out = 0
            do {
                if (i == 2) break
                out += i
                i += 1
            } while (true)
            println(out)
            println(i)
        }

