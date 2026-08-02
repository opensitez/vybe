// vybe-test: kotlin/control_flow/test_nested_when_and_for_interaction
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            val input = arrayOf(1, 2, 3, 4)
            var marks = ""
            for (value in input) {
                when {
                    value == 1 -> marks += "A"
                    value in 2..3 -> marks += "B"
                    else -> marks += "C"
                }
            }
            println(marks)
        }

