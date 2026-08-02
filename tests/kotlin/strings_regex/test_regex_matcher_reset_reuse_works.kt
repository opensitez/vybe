// vybe-test: kotlin/strings_regex/test_regex_matcher_reset_reuse_works
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val matcher = Regex("a(\\d)").toPattern().matcher("a1x")
            __check((matcher.find()).toString(), "true")
            __check((matcher.group(1)).toString(), "1")

            matcher.reset("a2")
            __check((matcher.find()).toString(), "true")
            __check((matcher.group(1)).toString(), "2")
        }
