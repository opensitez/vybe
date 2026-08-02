// vybe-test: kotlin/ranges/test_reversed_range_step_two
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            var result = ""
            for (value in 10 downTo 1 step 3) {
                result += value.toString()
            }
            println(result)
        }

