// vybe-test: kotlin/strings_regex/test_regex_replace_with_group_mapping
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val input = "id:42:done"
            val pattern = Regex("id:(\\d+):done")
            val output = pattern.replace(input) { match ->
                "ID=" + match.groups[1]!!.value
            }
            __check((output).toString(), "ID=42")
        }
