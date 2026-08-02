// vybe-test: kotlin/ranges/test_down_to_range_iteration
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            var output = ""
            for (i in 5 downTo 3) {
                output += i.toString()
            }
            println(output)
        }

