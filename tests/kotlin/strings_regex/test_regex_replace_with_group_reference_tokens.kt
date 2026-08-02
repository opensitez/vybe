// vybe-test: kotlin/strings_regex/test_regex_replace_with_group_reference_tokens
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("(\\w)(\\d)")
            __check((pattern.replace("a1 b2", "\$2-\$1")).toString(), "1-a 2-b")
        }
