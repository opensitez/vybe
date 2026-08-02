// vybe-test: kotlin/do_while_control/test_do_while_counter_with_return
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var i = 0
            var out = 0
            do {
                if (i == 3) {
                    out += 1
                    i = 10
                } else {
                    out += 2
                }
                i += 1
            } while (i < 5)
            println(out)
        }

