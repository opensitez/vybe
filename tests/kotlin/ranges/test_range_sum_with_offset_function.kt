// vybe-test: kotlin/ranges/test_range_sum_with_offset_function
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun span(start: Int, length: Int): IntRange {
            return start..(start + length)
        }

        fun main() {
            var total = 0
            for (value in span(2, 3)) {
                total += value
            }
            println(total)
        }

