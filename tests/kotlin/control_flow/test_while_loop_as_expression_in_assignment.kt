// vybe-test: kotlin/control_flow/test_while_loop_as_expression_in_assignment
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var i = 0
            val sum = run {
                var acc = 0
                while (i < 3) {
                    acc += i
                    i += 1
                }
                acc
            }
            println(sum)
        }

