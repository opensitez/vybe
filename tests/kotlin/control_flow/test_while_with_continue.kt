// vybe-test: kotlin/control_flow/test_while_with_continue
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var i = 0
            while (i < 5) {
                i += 1
                if (i % 2 == 0) continue
                println(i)
            }
        }

