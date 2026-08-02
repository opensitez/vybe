// vybe-test: kotlin/control_flow/test_labeled_continue_from_nested_for
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var out = ""
            outer@ for (row in 1..4) {
                for (col in 1..4) {
                    if (col == 3) continue@outer
                    out += "${row}${col}|"
                }
            }
            println(out)
        }

