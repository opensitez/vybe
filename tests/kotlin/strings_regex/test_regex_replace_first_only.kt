// vybe-test: kotlin/strings_regex/test_regex_replace_first_only
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("\\d+")
            __check((pattern.replaceFirst("a1b22c333", "#")).toString(), "a#b22c333")
            __check((pattern.replaceFirst("abc", "#")).toString(), "abc")
        }
