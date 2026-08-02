// vybe-test: kotlin/control_flow/test_for_loop_with_if_filter_in_body
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var values = ""
            for (i in 1..8) {
                if (i % 2 == 1) {
                    continue
                }
                values += i.toString()
            }
            println(values)
        }

