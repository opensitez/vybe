// vybe-test: kotlin/strings_regex/test_regex_from_literal_treats_as_plain_text
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex.fromLiteral("a+b")
            __check((pattern.containsMatchIn("c a+b d")).toString(), "true")
            __check((pattern.matches("a+b")).toString(), "true")
            __check((pattern.matches("aaab")).toString(), "false")
        }
