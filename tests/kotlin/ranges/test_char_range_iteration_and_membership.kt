// vybe-test: kotlin/ranges/test_char_range_iteration_and_membership
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            var output = ""
            for (value in 'a'..'d') {
                output += value.toString()
            }
            println(output)
            println('b' in 'a'..'d')
            println('x' in 'a'..'d')
        }

