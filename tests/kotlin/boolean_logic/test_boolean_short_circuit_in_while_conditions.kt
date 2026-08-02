// vybe-test: kotlin/boolean_logic/test_boolean_short_circuit_in_while_conditions
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun main() {
            var i = 0
            var steps = 0
            while (i < 3 && steps < 5) {
                steps++
                i++
            }
            println(i)
            println(steps)
            println(i == 3 && steps == 3)
        }

