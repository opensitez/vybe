// vybe-test: kotlin/control_flow/test_while_else_if_in_loop_condition
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var i = 0
            var evenCount = 0
            var oddCount = 0
            while (i < 6) {
                if (i % 2 == 0) {
                    evenCount += 1
                } else {
                    oddCount += 1
                }
                i += 1
            }
            println(evenCount)
            println(oddCount)
        }

