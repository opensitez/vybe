// vybe-test: kotlin/do_while_control/test_do_while_nested_break
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var i = 0
            do {
                i += 1
                if (i == 2) break
            } while (i < 4)
            println(i)
        }

