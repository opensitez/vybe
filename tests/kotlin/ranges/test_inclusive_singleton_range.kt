// vybe-test: kotlin/ranges/test_inclusive_singleton_range
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            var output = ""
            for (value in 3..3) {
                output += value.toString()
            }
            println(output)
            println(3 in 3..3)
            println(4 in 3..3)
        }

