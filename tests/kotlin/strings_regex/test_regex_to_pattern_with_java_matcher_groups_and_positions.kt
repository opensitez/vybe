// vybe-test: kotlin/strings_regex/test_regex_to_pattern_with_java_matcher_groups_and_positions
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun main() {
            val matcher = Regex("(\\w)(\\d)").toPattern().matcher("a1 b2 c3")
            var trace = ""
            while (matcher.find()) {
                trace += matcher.group(1)
                trace += matcher.group(2)
                trace += matcher.start().toString()
                trace += matcher.end().toString()
                trace += "|"
            }
            println(trace)
        }

