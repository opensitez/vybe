// vybe-test: kotlin/do_while_control/test_do_while_executes_once
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var i = 10
            var seen = 0
            do {
                seen += 1
            } while (i < 0)
            println(seen)
        }

