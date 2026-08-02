// vybe-test: kotlin/control_flow/test_for_while_do_when_combination
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var trace = ""
            for (i in 1..6) {
                when (i % 3) {
                    0 -> continue
                    1 -> trace += "a"
                    else -> trace += "b"
                }
                if (i > 4) break
            }
            println(trace)
        }

