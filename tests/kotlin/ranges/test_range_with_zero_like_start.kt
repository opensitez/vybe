// vybe-test: kotlin/ranges/test_range_with_zero_like_start
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            var output = ""
            for (value in 0..3) {
                output += value.toString()
            }
            println(output)
            println(0 in 0 until 3)
            println(3 in 0 until 3)
        }

