// vybe-test: kotlin/strings/test_substring_invalid_range_throws
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun main() {
            val word = "abc"
            try {
                println(word.substring(5))
            } catch (e: Exception) {
                println("error")
            }
        }

