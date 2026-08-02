// vybe-test: kotlin/operators/test_down_to_and_step_loop_operator
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun main() {
            var reversed = ""
            var sum = 0
            for (value in 7 downTo 3 step 2) {
                reversed += value.toString()
                sum += value
            }
            println(reversed)
            println(sum)
        }

