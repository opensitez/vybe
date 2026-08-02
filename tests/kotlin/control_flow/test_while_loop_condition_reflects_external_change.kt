// vybe-test: kotlin/control_flow/test_while_loop_condition_reflects_external_change
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

var done = false

        fun shouldRun(): Boolean {
            return !done
        }

        fun main() {
            var total = 0
            while (shouldRun()) {
                total += 1
                done = true
            }
            println(total)
            println(shouldRun())
        }

