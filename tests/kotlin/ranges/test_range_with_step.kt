// vybe-test: kotlin/ranges/test_range_with_step
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            var total = 0
            for (i in 1..10 step 3) {
                total += i
            }
            println(total)
        }

