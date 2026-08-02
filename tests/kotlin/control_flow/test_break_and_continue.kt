// vybe-test: kotlin/control_flow/test_break_and_continue
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            for (i in 1..5) {
                if (i == 2) continue
                if (i == 4) break
                println(i)
            }
        }

