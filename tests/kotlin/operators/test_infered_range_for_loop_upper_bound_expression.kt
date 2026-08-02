// vybe-test: kotlin/operators/test_infered_range_for_loop_upper_bound_expression
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun main() {
            var result = 0
            val multiplier = 1
            for (value in 1..(2 + multiplier)) {
                result += value
            }
            println(result)
        }

