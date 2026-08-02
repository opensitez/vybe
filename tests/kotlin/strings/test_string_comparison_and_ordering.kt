// vybe-test: kotlin/strings/test_string_comparison_and_ordering
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("abc" < "abd").toString(), "true")
            __check(("abc" == "abc").toString(), "true")
            __check(("ABC" < "abc").toString(), "true")
            __check(("ABC".equals("abc", ignoreCase = true)).toString(), "true")
        }
