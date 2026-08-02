// vybe-test: kotlin/control_flow/test_do_while_condition_uses_updated_variable
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var x = 0
            do {
                println(x)
                x += 3
            } while (x < 8)
        }

