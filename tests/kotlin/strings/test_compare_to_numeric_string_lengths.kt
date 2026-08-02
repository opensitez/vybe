// vybe-test: kotlin/strings/test_compare_to_numeric_string_lengths
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun main() {
            val words = listOf("a", "ab", "abc")
            var shortest = words[0]
            for (word in words) {
                if (word.length < shortest.length) {
                    shortest = word
                }
            }
            println(shortest)
        }

