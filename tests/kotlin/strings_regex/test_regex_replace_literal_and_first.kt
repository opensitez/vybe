// vybe-test: kotlin/strings_regex/test_regex_replace_literal_and_first
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "a+b+c"
            val escaped = Regex.escape(value)
            __check((escaped).toString(), "\\Qa+b+c\\E")
            __check((Regex(escaped).replace(value, "_")).toString(), "_")
            __check((Regex("\\d+").replaceFirst("x1x2x", "NUM")).toString(), "xNUMx2x")
        }
