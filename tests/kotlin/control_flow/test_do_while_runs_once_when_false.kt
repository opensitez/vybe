// vybe-test: kotlin/control_flow/test_do_while_runs_once_when_false
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var n = 0
            do {
                n += 1
            } while (n > 10)
            println(n)
        }

