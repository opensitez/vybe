// vybe-test: kotlin/string_search_apis/test_contains_with_ignore_case
// origin: languages/kotlin/tests/kotlin/test_string_search_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "Kotlin"
            __check((text.contains("kin")).toString(), "true")
            __check((text.contains("KIN", ignoreCase = true)).toString(), "true")
            __check((text.contains('K')).toString(), "true")
            __check((text.contains('z')).toString(), "false")
        }
