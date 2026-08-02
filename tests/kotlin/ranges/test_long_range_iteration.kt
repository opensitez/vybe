// vybe-test: kotlin/ranges/test_long_range_iteration
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            var total = 0L
            for (value in 1L..7L step 2) {
                total += value
            }
            println(total)
            println(6L in 1L..7L)
            println(8L in 1L until 8L)
        }

