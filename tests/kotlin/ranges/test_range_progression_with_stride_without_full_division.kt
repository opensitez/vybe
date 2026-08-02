// vybe-test: kotlin/ranges/test_range_progression_with_stride_without_full_division
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            var values = ""
            for (value in 0..11 step 4) {
                values += value.toString()
                if (value > 8) {
                    values += "|"
                }
            }
            println(values)
            println(10 in (0..11 step 4))
            println(8 in (0..11 step 4))
        }

