// vybe-test: kotlin/control_flow/test_nested_loops
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            for (i in 1..2) {
                for (j in 1..2) {
                    println(i * 10 + j)
                }
            }
        }

