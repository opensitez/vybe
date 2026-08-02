// vybe-test: kotlin/control_flow/test_while_with_nested_break
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var outer = 0
            while (outer < 3) {
                var inner = 0
                while (inner < 3) {
                    if (inner == 1) break
                    println(outer * 10 + inner)
                    inner += 1
                }
                outer += 1
            }
        }

