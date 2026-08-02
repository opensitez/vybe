// vybe-test: kotlin/control_flow/test_if_guard_in_when
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            val x = 10
            when {
                x < 0 -> println("negative")
                x == 0 -> println("zero")
                x > 0 -> println("positive")
            }
        }

