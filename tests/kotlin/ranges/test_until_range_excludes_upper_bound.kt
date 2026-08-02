// vybe-test: kotlin/ranges/test_until_range_excludes_upper_bound
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            var total = 0
            for (i in 1 until 4) {
                total += i
            }
            println(total)
        }

