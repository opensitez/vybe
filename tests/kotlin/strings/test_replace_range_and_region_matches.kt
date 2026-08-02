// vybe-test: kotlin/strings/test_replace_range_and_region_matches
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "abcdef"
            __check((value.replaceRange(1, 3, "ZZ")).toString(), "aZZdef")
            __check((value.regionMatches(1, "CD", 0, 2, ignoreCase = true)).toString(), "true")
            __check((value.regionMatches(1, "Cd", 0, 2, ignoreCase = true)).toString(), "false")
        }
