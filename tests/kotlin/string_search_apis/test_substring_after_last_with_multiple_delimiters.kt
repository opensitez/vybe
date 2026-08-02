// vybe-test: kotlin/string_search_apis/test_substring_after_last_with_multiple_delimiters
// origin: languages/kotlin/tests/kotlin/test_string_search_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "/home/user/docs/readme.txt"
            __check((text.substringAfterLast("/")).toString(), "readme.txt")
            __check((text.substringBeforeLast("/")).toString(), "/home/user/docs")
            __check(("abc".substringAfterLast("/", "na")).toString(), "na")
        }
