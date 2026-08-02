// vybe-test: kotlin/control_flow/test_while_with_continue_and_nested_if
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var count = 0
            var i = 0
            while (i < 6) {
                i += 1
                if (i == 2 || i == 5) {
                    continue
                }
                count += i
            }
            println(count)
        }

