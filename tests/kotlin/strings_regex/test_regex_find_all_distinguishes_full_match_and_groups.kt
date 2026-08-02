// vybe-test: kotlin/strings_regex/test_regex_find_all_distinguishes_full_match_and_groups
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pattern = Regex("(\\w)-(\\d)")
            val first = pattern.find("a-1")
            __check((first?.groupValues?.getOrNull(0) ?: "none").toString(), "a-1")
            __check((first?.groupValues?.getOrNull(1) ?: "none").toString(), "a")
            __check((first?.groupValues?.getOrNull(2) ?: "none").toString(), "1")
            __check((first?.groupValues?.size ?: 0).toString(), "3")
        }
