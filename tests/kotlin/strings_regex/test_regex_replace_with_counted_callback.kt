// vybe-test: kotlin/strings_regex/test_regex_replace_with_counted_callback
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var index = 0
            val pattern = Regex("\\d")
            val output = pattern.replace("a1b2c3") { match ->
                val value = "${index}:${match.value}"
                index += 1
                value
            }
            __check((output).toString(), "a0:1b1:2c2:3")
            __check((index).toString(), "3")
        }
