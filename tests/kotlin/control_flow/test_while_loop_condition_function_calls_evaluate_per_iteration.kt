// vybe-test: kotlin/control_flow/test_while_loop_condition_function_calls_evaluate_per_iteration
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

var calls = 0

        fun shouldContinue(): Boolean {
            calls += 1
            return calls < 3
        }

        fun main() {
            var count = 0
            while (shouldContinue()) {
                count += 1
            }
            println(count)
            println(calls)
        }

