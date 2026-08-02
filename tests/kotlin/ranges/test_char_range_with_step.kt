// vybe-test: kotlin/ranges/test_char_range_with_step
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            var output = ""
            for (value in 'a'..'f' step 2) {
                output += value.toString()
            }
            println(output)
            println('e' in 'a'..'f' step 2)
            println('d' in 'a'..'f' step 2)
        }

