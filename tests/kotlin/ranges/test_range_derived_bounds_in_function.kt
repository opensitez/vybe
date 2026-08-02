// vybe-test: kotlin/ranges/test_range_derived_bounds_in_function
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun sumInRange(start: Int, end: Int): Int {
            var total = 0
            for (value in start..end) {
                total += value
            }
            return total
        }

        fun main() {
            println(sumInRange(3, 5))
        }

