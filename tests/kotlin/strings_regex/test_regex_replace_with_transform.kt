// vybe-test: kotlin/strings_regex/test_regex_replace_with_transform
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("(\\d+)")
            val out = pattern.replace("a12b34c") { match -> match.value.reversed() }
            __check((out).toString(), "a21b43c")
        }
