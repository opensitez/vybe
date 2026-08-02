// vybe-test: kotlin/do_while_control/test_do_while_boolean_expression
// origin: languages/kotlin/tests/kotlin/test_do_while_control.rs

fun main() {
            var i = 0
            var out = 0
            do {
                out += if (i % 2 == 0) i else 0
                i += 1
            } while (i < 6)
            println(out)
        }

