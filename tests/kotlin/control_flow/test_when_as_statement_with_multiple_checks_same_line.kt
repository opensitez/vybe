// vybe-test: kotlin/control_flow/test_when_as_statement_with_multiple_checks_same_line
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            val x = 5
            when (x) {
                in 1..3, in 7..9 -> println("edge")
                5, 6 -> println("middle")
                else -> println("other")
            }
        }

