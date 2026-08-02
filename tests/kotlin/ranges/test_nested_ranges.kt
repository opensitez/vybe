// vybe-test: kotlin/ranges/test_nested_ranges
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            var count = 0
            for (row in 1..2) {
                for (col in row..3) {
                    count += col
                }
            }
            println(count)
        }

