// vybe-test: kotlin/ranges/test_descending_range_with_empty_bounds
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            var count = 0
            for (value in 2 downTo 5) {
                count += value
            }
            println(count)
            println(5 in 2 downTo 5)
        }

