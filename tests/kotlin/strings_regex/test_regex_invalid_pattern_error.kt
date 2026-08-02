// vybe-test: kotlin/strings_regex/test_regex_invalid_pattern_error
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun main() {
            try {
                val pattern = Regex("[")
                println(pattern)
            } catch (e: java.lang.RuntimeException) {
                println("bad")
            }
        }

