// vybe-test: kotlin/control_flow/test_while_with_break
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var count = 0
            while (true) {
                if (count == 3) break
                println(count)
                count += 1
            }
        }

