// vybe-test: kotlin/strings_regex/test_regex_matcher_with_all_occurrences_and_matcher_state
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun main() {
            val pattern = Regex("\\b\\w+\\b")
            val matcher = pattern.toPattern().matcher("one two three")
            var matches = ""
            while (matcher.find()) {
                matches += matcher.group()
                matches += ":"
                matches += matcher.start().toString()
                matches += "-"
                matches += matcher.end().toString()
                matches += ";"
            }
            println(matches)
        }

