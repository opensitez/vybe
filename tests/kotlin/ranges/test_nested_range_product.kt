// vybe-test: kotlin/ranges/test_nested_range_product
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            var total = 0
            for (a in 1..3) {
                for (b in a..4) {
                    total += a * b
                }
            }
            println(total)
        }

