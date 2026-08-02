// vybe-test: kotlin/string_search_apis/test_substring_before_after_variants
// origin: languages/kotlin/tests/kotlin/test_string_search_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "a=b=c"
            __check((text.substringAfter("=")).toString(), "b=c")
            __check((text.substringAfter("z", "none")).toString(), "none")
            __check((text.substringBefore("=")).toString(), "a")
            __check((text.substringBeforeLast("=")).toString(), "a=b")
            __check((text.substringAfterLast("=")).toString(), "c")
        }
