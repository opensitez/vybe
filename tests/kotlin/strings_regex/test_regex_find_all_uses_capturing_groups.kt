// vybe-test: kotlin/strings_regex/test_regex_find_all_uses_capturing_groups
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun main() {
            val pattern = Regex("(\\w+)(\\d)")
            val matches = pattern.findAll("a1 b22 c3")
            var output = ""
            for (item in matches) {
                output += item.groupValues[1]
                output += "-"
                output += item.groupValues[2]
                output += ";"
            }
            println(output)
        }

