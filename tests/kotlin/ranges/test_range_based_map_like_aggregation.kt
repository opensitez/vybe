// vybe-test: kotlin/ranges/test_range_based_map_like_aggregation
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun build(start: Int, end: Int, step: Int): Int {
            var total = 1
            for (value in start..end step step) {
                total *= value
            }
            return total
        }

        fun main() {
            println(build(1, 4, 1))
            println(build(2, 6, 2))
        }

