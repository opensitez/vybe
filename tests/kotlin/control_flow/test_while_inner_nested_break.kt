// vybe-test: kotlin/control_flow/test_while_inner_nested_break
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var outer = 0
            while (outer < 4) {
                var inner = 0
                while (inner < 4) {
                    inner += 1
                    if (inner == 3) {
                        break
                    }
                }
                outer += 1
                println(outer)
                if (outer == 2) {
                    break
                }
            }
        }

