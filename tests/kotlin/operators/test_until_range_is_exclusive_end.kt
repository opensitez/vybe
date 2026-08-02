// vybe-test: kotlin/operators/test_until_range_is_exclusive_end
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun main() {
            var total = 0
            for (value in 1 until 4) {
                total += value
            }
            println(total)
        }

