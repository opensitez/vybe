// vybe-test: kotlin/control_flow/test_while_with_if_assignment
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var total = 0
            var i = 0
            while (i < 4) {
                val next = if (i % 2 == 0) i else i + 1
                total += next
                i += 1
            }
            println(total)
        }

