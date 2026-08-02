// vybe-test: kotlin/string_search_apis/test_starts_with_end_with
// origin: languages/kotlin/tests/kotlin/test_string_search_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "KotlinLang"
            __check((text.startsWith("Kot")).toString(), "true")
            __check((text.startsWith("tin", 2)).toString(), "false")
            __check((text.endsWith("Lang")).toString(), "true")
            __check((text.endsWith("lang", ignoreCase = true)).toString(), "true")
        }
