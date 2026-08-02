// vybe-test: kotlin/ranges/test_reversed_range_iteration_matches_original_elements
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            var ascending = ""
            for (value in (5 downTo 1).reversed()) {
                ascending += value.toString()
            }
            println(ascending)
            println((5 downTo 1).contains(3))
            println((5 downTo 1).contains(0))
        }

