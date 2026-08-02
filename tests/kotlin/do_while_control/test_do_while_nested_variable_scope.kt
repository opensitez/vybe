// vybe-test: kotlin/do_while_control/test_do_while_nested_variable_scope
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var out = 0
            do {
                var inner = 0
                inner += 1
                out += inner
            } while (out < 3)
            println(out)
        }

