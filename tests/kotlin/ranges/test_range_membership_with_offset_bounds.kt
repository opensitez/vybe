// vybe-test: kotlin/ranges/test_range_membership_with_offset_bounds
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            val lower = 2
            val upper = 5
            var result = ""
            for (value in (lower + 1)..upper) {
                result += value.toString()
            }
            println(result)
            println((lower + 1) in (lower + 1)..upper)
            println((upper) in (lower + 1)..upper)
        }

