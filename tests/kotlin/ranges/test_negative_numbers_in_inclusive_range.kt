// vybe-test: kotlin/ranges/test_negative_numbers_in_inclusive_range
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            var total = 0
            for (value in -2..2) {
                total += value
            }
            println(total)
            println(-1 in -2..2)
            println(3 in -2..2)
        }

