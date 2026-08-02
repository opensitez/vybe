// vybe-test: kotlin/ranges/test_range_with_step_one_is_equivalent_to_plain_range
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            var a = ""
            for (value in 2..8 step 1) {
                a += value.toString()
            }
            var b = ""
            for (value in 2..8) {
                b += value.toString()
            }
            println(a)
            println(b)
            println(a == b)
        }

